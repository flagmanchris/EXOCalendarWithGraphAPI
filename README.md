# EXOCalendarWithGraphAPI

Snippets required to setup a runbook on an azure automation account which removes cancelled meetings where a specific service account is the organiser

## 1) Add-MSIGraphPermissions.ps1
Use this an an example for granting Graph API permissions to an Azure Automation Account MSI. Required as a pre-req for the runbook code.

## 2) Remove-CancelledMeetings.ps1
Runbook for removing cancelled meetings from specific mailboxes organised by a specific account
