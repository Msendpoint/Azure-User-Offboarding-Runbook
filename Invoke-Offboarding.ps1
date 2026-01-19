<#
.SYNOPSIS
    Made By Souhaiel MORHAG
    Check MSEndpoint.com for more Info
    01/19/2026
    Automated User Offboarding with Governance & Approval Tracking.
    Handles forwarding, shared mailbox conversion, permissions, OOF, and compliance logging.

.DESCRIPTION
    This script is designed to be run as an Azure Automation Runbook (PowerShell 7.2).
    It performs the following actions in order:
    1. Authenticates via Managed Identity.
    2. Converts the source User Mailbox to a Shared Mailbox (optional).
    3. Grants 'Full Access' and 'Send As' permissions to target delegates.
    4. Configures Email Forwarding (Internal/External) and toggles 'Keep Copy'.
    5. Sets an Automatic Reply (Out of Office) with a dynamic template.
    6. Logs the operation AND Governance data (Ticket Ref, Proof URL) to Azure Table Storage.
    7. Sends a detailed notification email with links to the proof.

.PARAMETER ApprovalTicketRef
    REQUIRED. The Service Management ticket number authorizing this request (e.g., "INC0012345", "RITM9999").
    
.PARAMETER ApprovalDocLink
    OPTIONAL. A URL to the signed PDF/Email approval (SharePoint link or Blob URL) for audit.

.PARAMETER SourceUserEmail
    The email of the departing user.

.PARAMETER TargetUserEmail
    The primary email for forwarding (User Y).

.PARAMETER AdditionalDelegates
    Comma-separated list of OTHER emails that need permissions.

.PARAMETER ConvertToShared
    Bool. If true, converts to Shared Mailbox.

.PARAMETER DeliverToMailboxAndForward
    Bool. If true, keeps a copy of the email in the source mailbox.

.PARAMETER GrantFullAccess
    Bool. Grant 'Full Access' (Read/Manage).

.PARAMETER GrantSendAs
    Bool. Grant 'Send As' rights.

.PARAMETER OofTemplate
    String. "LeftCompany", "LongLeave", "None".

.PARAMETER OofEndDate
    DateTime. When to stop the auto-reply (usually same as forwarding duration).

.PARAMETER AuditStorageAccountName
    Azure Storage Account for logging.

.PARAMETER NotificationEmail
    Who receives the final report.
#>

param(
    # --- GOVERNANCE PARAMETERS (NEW) ---
    [Parameter(Mandatory=$true)]
    [string]$ApprovalTicketRef,  # e.g., INC123456

    [Parameter(Mandatory=$false)]
    [string]$ApprovalDocLink,    # e.g., https://sharepoint.../approval.pdf

    # --- TECHNICAL PARAMETERS ---
    [Parameter(Mandatory=$true)]
    [string]$SourceUserEmail,

    [Parameter(Mandatory=$true)]
    [string]$TargetUserEmail,

    [Parameter(Mandatory=$false)]
    [string]$AdditionalDelegates, 

    [Parameter(Mandatory=$true)]
    [bool]$ConvertToShared,

    [Parameter(Mandatory=$true)]
    [bool]$DeliverToMailboxAndForward,

    [Parameter(Mandatory=$true)]
    [bool]$GrantFullAccess,

    [Parameter(Mandatory=$true)]
    [bool]$GrantSendAs,

    [Parameter(Mandatory=$true)]
    [ValidateSet("LeftCompany", "LongLeave", "None")]
    [string]$OofTemplate,

    [Parameter(Mandatory=$false)]
    [DateTime]$OofEndDate = (Get-Date).AddMonths(3),

    [Parameter(Mandatory=$true)]
    [string]$AuditStorageAccountName,

    [Parameter(Mandatory=$true)]
    [string]$NotificationEmail
)

# --- CONFIGURATION (UPDATE THESE BEFORE DEPLOYING) ---
$OrganizationDomain = "yourtenant.onmicrosoft.com"  # TODO: Replace with your Tenant Domain
$StorageResourceGroup = "rg-automation-prod"        # TODO: Replace with your Storage Account's Resource Group
$SenderEmailAddress   = "noreply@yourdomain.com"    # TODO: Replace with a valid sender email for notifications
# -----------------------------------------------------

# 1. AUTHENTICATION
try {
    Write-Output "Connecting to Azure & Exchange Online..."
    Connect-AzAccount -Identity | Out-Null
    Connect-ExchangeOnline -ManagedIdentity -Organization $OrganizationDomain | Out-Null
}
catch {
    Write-Error "Authentication failed. Ensure System Assigned Managed Identity is enabled and has 'Exchange.ManageAsApp' permissions."
    throw $_
}

$logEntries = @()
$logEntries += "Start Time: $(Get-Date)"
$logEntries += "<strong>Governance:</strong> Ticket Ref: $ApprovalTicketRef"

# 2. CONVERSION (TO SHARED)
if ($ConvertToShared) {
    try {
        Write-Output "Converting $SourceUserEmail to Shared Mailbox..."
        Set-Mailbox -Identity $SourceUserEmail -Type Shared -ErrorAction Stop
        $logEntries += "SUCCESS: Converted to Shared Mailbox."
    }
    catch {
        $logEntries += "ERROR: Conversion failed ($($_))"
    }
}

# 3. PERMISSIONS (DELEGATION)
$usersToGrant = @($TargetUserEmail)
if (-not [string]::IsNullOrWhiteSpace($AdditionalDelegates)) {
    $usersToGrant += $AdditionalDelegates -split ","
}

foreach ($user in $usersToGrant) {
    $user = $user.Trim()
    
    # Grant Full Access
    if ($GrantFullAccess) {
        try {
            Add-MailboxPermission -Identity $SourceUserEmail -User $user -AccessRights FullAccess -InheritanceType All -AutoMapping $false -ErrorAction Stop
            $logEntries += "SUCCESS: Granted Full Access to $user"
        } catch { $logEntries += "ERROR: Failed Full Access for $user ($($_))" }
    }

    # Grant Send As
    if ($GrantSendAs) {
        try {
            Add-RecipientPermission -Identity $SourceUserEmail -Trustee $user -AccessRights SendAs -Confirm:$false -ErrorAction Stop
            $logEntries += "SUCCESS: Granted Send As to $user"
        } catch { $logEntries += "ERROR: Failed Send As for $user ($($_))" }
    }
}

# 4. FORWARDING
try {
    Write-Output "Setting forwarding to $TargetUserEmail..."
    Set-Mailbox -Identity $SourceUserEmail -ForwardingSmtpAddress $TargetUserEmail -DeliverToMailboxAndForward $DeliverToMailboxAndForward -ErrorAction Stop
    $logEntries += "SUCCESS: Forwarding set to $TargetUserEmail. Keep Copy: $DeliverToMailboxAndForward"
}
catch {
    $logEntries += "CRITICAL ERROR: Forwarding failed ($($_))"
}

# 5. AUTOMATIC REPLIES (OOF)
if ($OofTemplate -ne "None") {
    $oofMessage = ""
    switch ($OofTemplate) {
        "LeftCompany" {
            $oofMessage = "<html><body><p>Bonjour / Hello,</p><p>Je ne fais plus partie de la société (I have left the company).</p><p>Merci de contacter <b>$TargetUserEmail</b> pour toute demande.</p></body></html>"
        }
        "LongLeave" {
            $oofMessage = "<html><body><p>Bonjour,</p><p>Je suis absent jusqu'au $($OofEndDate.ToString('dd/MM/yyyy')).</p><p>En cas d'urgence, contactez <b>$TargetUserEmail</b>.</p></body></html>"
        }
    }
    try {
        Set-MailboxAutoReplyConfiguration -Identity $SourceUserEmail -AutoReplyState Scheduled -StartTime (Get-Date) -EndTime $OofEndDate -InternalMessage $oofMessage -ExternalMessage $oofMessage -ErrorAction Stop
        $logEntries += "SUCCESS: OOF set ($OofTemplate) until $OofEndDate"
    }
    catch {
        $logEntries += "ERROR: OOF configuration failed ($($_))"
    }
}

# 6. AUDIT LOGGING (AZURE TABLE)
$tableName = "ForwardingTracking"
try {
    $ctx = (Get-AzStorageAccount -ResourceGroupName $StorageResourceGroup -Name $AuditStorageAccountName).Context
    New-AzStorageTable -Name $tableName -Context $ctx -ErrorAction SilentlyContinue
    
    # PartitionKey = Domain, RowKey = UserEmail (Sanitized)
    $rowKey = $SourceUserEmail.Replace("@","_")
    
    Add-AzTableEntity -Table $tableName -Context $ctx -Entity @{
        "PartitionKey" = $SourceUserEmail.Split('@')[1];
        "RowKey"       = $rowKey;
        "UserEmail"    = $SourceUserEmail;
        "Action"       = "DisableForwarding";
        "ExpiryDate"   = $OofEndDate;
        "Status"       = "Active";
        "TicketRef"    = $ApprovalTicketRef;
        "ProofUrl"     = $ApprovalDocLink
    } -ErrorAction Ignore
    $logEntries += "SUCCESS: Audit logged to Table Storage (Ticket: $ApprovalTicketRef). Cleanup scheduled on $OofEndDate."
}
catch {
    $logEntries += "WARNING: Failed to write to Azure Table. Cleanup might fail. Error: $_"
}

# 7. REPORTING (EMAIL NOTIFICATION)
$proofHtml = ""
if (-not [string]::IsNullOrEmpty($ApprovalDocLink)) {
    $proofHtml = "<p><strong>Approval Attachment:</strong> <a href='$ApprovalDocLink'>View Authorized Proof</a></p>"
}

$emailSubject = "Offboarding Complete: $SourceUserEmail (Ticket $ApprovalTicketRef)"
$finalReportLogs = $logEntries -join "<br>"

$emailBodyHtml = @"
<h2>Offboarding Execution Report</h2>
<div style='background-color:#f0f0f0; padding:15px; border-left: 5px solid #0078D4; margin-bottom: 20px;'>
    <h3 style='margin-top:0;'>Governance & Authorization</h3>
    <p><strong>ServiceNow Ticket:</strong> $ApprovalTicketRef</p>
    $proofHtml
    <p><strong>Executed By:</strong> Azure Automation (Managed Identity)</p>
    <p><strong>Expiry Date:</strong> $OofEndDate</p>
</div>
<hr>
<h3>Technical Execution Log</h3>
<p>$finalReportLogs</p>
"@

$message = @{
    Subject = $emailSubject
    Body    = @{ ContentType = "HTML"; Content = $emailBodyHtml }
    ToRecipients = @(@{ EmailAddress = @{ Address = $NotificationEmail } })
}

try {
    Send-MgUserMail -UserId $SenderEmailAddress -BodyParameter $message
    Write-Output "Report sent to $NotificationEmail"
}
catch {
    Write-Error "Failed to send report email via Graph. Ensure 'Mail.Send' permission. Error: $_"
}
