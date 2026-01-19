# Azure Office 365 Offboarding Automation

This repository contains a PowerShell script designed for **Azure Automation Runbooks**. It automates the offboarding process for Office 365 users while maintaining strict compliance (GDPR, ISO 27001) regarding data access, retention, and governance.

## Features

* ✅ **Governance & Audit Proof:** Links every execution to a ServiceNow Ticket ID and an optional Proof URL (SharePoint/Blob).
* ✅ **Mailbox Conversion:** Converts User Mailbox to Shared Mailbox (optional).
* ✅ **Forwarding:** Sets up email forwarding to a target user (Internal/External).
* ✅ **Permissions:** Grants *Full Access* and *Send As* rights to delegates.
* ✅ **Auto-Replies (OOF):** Configures standardized Out-of-Office messages ("Left the company").
* ✅ **Compliance:** Limits forwarding duration (e.g., 3 months) and logs the expiration date.
* ✅ **Audit Trail:** Logs actions to **Azure Table Storage** for tracking.
* ✅ **Reporting:** Sends a completion report via email (ServiceNow compatible).

## Prerequisites

### 1. Azure Automation Account
* Create an Azure Automation Account.
* Enable **System Assigned Managed Identity**.

### 2. Required Modules
Import the following modules in your Automation Account (Runtime 7.2):
1.  `Az.Accounts`
2.  `Az.Storage`
3.  `ExchangeOnlineManagement`
4.  `Microsoft.Graph.Authentication`
5.  `Microsoft.Graph.Users.Actions`

### 3. Permissions (Managed Identity)
The Managed Identity requires the following permissions to function. Run these in Azure Cloud Shell:

| API | Permission | Type | Reason |
| :--- | :--- | :--- | :--- |
| **Exchange Online** | `Exchange.ManageAsApp` | App Role | To run `Set-Mailbox` without user login. |
| **Azure AD** | `Exchange Administrator` | Role | Admin access to modify mailboxes. |
| **Microsoft Graph** | `Mail.Send` | App Role | To send the completion report email. |
| **Microsoft Graph** | `User.Read.All` | App Role | To look up user object IDs. |

## Usage

### 1. Deploy the Script
1.  Create a new Runbook (PowerShell 7.2).
2.  Paste the content of `Invoke-Offboarding.ps1`.
3.  **Update variables:** Edit lines **104-106** to match your Tenant ID (`$OrganizationDomain`), Resource Group (`$StorageResourceGroup`), and Sender Email (`$SenderEmailAddress`).
4.  Publish the Runbook.

### 2. Run the Runbook
Input parameters explanation:

#### Governance (New)
* **ApprovalTicketRef:** (Required) The ServiceNow Incident or Request number (e.g., `RITM123456`) that authorizes this action.
* **ApprovalDocLink:** (Optional) A direct link to the approved PDF, email, or SharePoint document.

#### Technical
* **SourceUserEmail:** The departing user (e.g., `ex-employee@contoso.com`).
* **TargetUserEmail:** The receiver (e.g., `manager@contoso.com`).
* **ConvertToShared:** `true` to convert to Shared Mailbox first.
* **DeliverToMailboxAndForward:** `true` to keep a copy in the original mailbox.
* **GrantFullAccess / GrantSendAs:** `true` to give permissions to the Target User.
* **OofTemplate:** Select `LeftCompany` for a standard exit message.
* **NotificationEmail:** Where to send the report (e.g., `servicenow@contoso.com` or IT Admin).

## Compliance & Security

This script supports compliance with:
* **GDPR (Art 5.1.e):** Forwarding is temporary (default 3 months), adhering to storage limitation principles.
* **ISO 27001 (A.8.1.3):** Automates the acceptable use policy for asset transfer.
* **SOC 2 (CC5.1):** Provides an immutable audit log in Azure Table Storage linked to a specific authorization ticket.

---
*Disclaimer: This script is provided as-is. Please test in a non-production environment before deploying.*
