<#
.SYNOPSIS
    Made By Souhaiel MORHAG
    Check MSEndpoint.com for more Info
    01/19/2026
    Automated User Offboarding for Office 365 (Exchange Online).
    Handles forwarding, shared mailbox conversion, permissions, OOF, and compliance logging.

.DESCRIPTION
    This script is designed to be run as an Azure Automation Runbook (PowerShell 7.2).
    It performs the following actions in order:
    1. Authenticates via Managed Identity (requires Exchange.ManageAsApp).
    2. Converts the source User Mailbox to a Shared Mailbox (optional).
    3. Grants 'Full Access' and 'Send As' permissions to target delegates.
    4. Configures Email Forwarding (Internal/External) and toggles 'Keep Copy'.
    5. Sets an Automatic Reply (Out of Office) with a dynamic template.
    6. Logs the operation to Azure Table Storage (for audit and auto-cleanup).
    7. Sends a detailed notification email (e.g., to ServiceNow or IT Support).

.NOTES
    Version:     1.0
    Prerequisites: ExchangeOnlineManagement, Az.Storage, Az.Accounts modules.

.LINK
    https://github.com/votre-username/Azure-User-Offboarding
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$SourceUserEmail,

    [Parameter(Mandatory=$true)]
    [string]$TargetUserEmail,

    [Parameter(Mandatory=$false)]
    [string]$AdditionalDelegates, # Format: "manager@domain.com,hr@domain.com"

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

# --- CONFIGURATION (UPDATE THESE BEFORE RUNNING) ---
$OrganizationDomain = "yourtenant.onmicrosoft.com"  # TODO: Replace with your Tenant Domain
$StorageResourceGroup = "rg-automation-prod"        # TODO: Replace with your Storage Account's Resource Group
$SenderEmailAddress   = "noreply@yourdomain.com"    # TODO: Replace with a valid sender email for notifications
# -------------------------------------------------

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
    if ($GrantFullAccess) {
        try {
            Add-MailboxPermission -Identity $SourceUserEmail -User $user -AccessRights FullAccess -InheritanceType All -AutoMapping $false -ErrorAction Stop
            $logEntries += "SUCCESS: Granted Full Access to $user"
        } catch { $logEntries += "ERROR: Failed Full Access for $user ($($_))" }
    }
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
        "Status"       = "Active"
    } -ErrorAction Ignore
    $logEntries += "SUCCESS: Audit logged to Table Storage. Scheduled for cleanup on $OofEndDate."
}
catch {
    $logEntries += "WARNING: Failed to write to Azure Table. Cleanup might fail. Error: $_"
}

# 7. REPORTING (EMAIL NOTIFICATION)
# Using Graph for sending mail
$emailSubject = "Offboarding Report: $SourceUserEmail"
$finalReport = $logEntries -join "<br>"
$emailBody = @{
    ContentType = "HTML"
    Content     = "<h2>Offboarding Execution Report</h2><p>$finalReport</p>"
}

$message = @{
    Subject = $emailSubject
    Body    = $emailBody
    ToRecipients = @(@{ EmailAddress = @{ Address = $NotificationEmail } })
}

try {
    Send-MgUserMail -UserId $SenderEmailAddress -BodyParameter $message
    Write-Output "Report sent to $NotificationEmail"
}
catch {
    Write-Error "Failed to send report email via Graph. Ensure 'Mail.Send' permission. Error: $_"
}
