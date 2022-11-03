<#Graph API permissions to add
"Calendars.ReadWrite" for reading/removing the meetings
"Mail.Send" for sending a summary email if needed
"User.Read.All" for getting the list of users to work with if needed
#>
$permissionsToAdd = @(
    "Calendars.ReadWrite"
    "Mail.Send"
    "User.Read.All"
)

#Graph API App ID (Same for all tenants)
$GraphAppId = "00000003-0000-0000-c000-000000000000"

Import-Module AzureAD
Connect-AzureAD

#region get the Managed Service Identity App details with either name or ID
#Search with APP Name 
$mSIDisplayName = "AUTOMATION ACCOUNT NAME"
$mSI = Get-AzureADServicePrincipal -Filter "displayName eq '$mSIDisplayName'"
#Or use the AppID from the portal 
$mSIID = "ATUOMATION ACCOUNT IDENTITY"
$mSI = Get-AzureADServicePrincipal -ObjectId $mSIID
#endregion

#get the Graph API App 
$graphSP = Get-AzureADServicePrincipal -Filter "appId eq '$GraphAppId'"

foreach ($perm in $permissionsToAdd) {
    #Get the app role from Graph API app
    $appRole = $graphSP.AppRoles | Where-Object {$_.Value -eq $perm -and $_.AllowedMemberTypes -contains "Application"}
    #Add the app role to the Managed Service Identity
    New-AzureAdServiceAppRoleAssignment -ObjectId $mSI.ObjectId -PrincipalId $mSI.ObjectId -ResourceId $graphSP.ObjectId -Id $appRole.Id
}
DisConnect-AzureAD
