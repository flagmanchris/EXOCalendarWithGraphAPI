<#
  Example for a runbook running in an Azure automation account; using the managed service identity for authentication. MSI will need Graph API permissions assigning.
#>
#region variables and initialise stuff
#Just used to ensure not running against more users than expected 
$maxUserCount = 150
$stopwatch =  [system.diagnostics.stopwatch]::StartNew()
try {
    Import-Module Microsoft.Graph.Authentication, Microsoft.Graph.Calendar, Microsoft.Graph.Users, Microsoft.Graph.Users.Actions
}
catch {
    Write-Error "Failed to import modules"
}

try {
    $connect = Connect-AzAccount -Identity
    Write-Output "Connected to Azure"
    $accessToken = (Get-AzAccessToken -ResourceTypeName MSGraph).token
    Write-Output "Acquired token"
    $date = get-date -hour 0 -minute 0 -Second 0 -Millisecond 0
    #filter for events after today, cancelled and from service account. Uses the service account displayname
    $filter = "Organizer/emailAddress/Name eq 'Account, Service' and start/DateTime ge '$date' and iscancelled eq true" 
}
catch {
    Write-Error "Error during initialisation"
}
#endregion 

#region connect to graph
try {
    $session = Connect-MgGraph -AccessToken $accessToken
    Write-Output "Connected to Graph API"
}
catch {
    Write-Error "Failed to connect to Graph API"
}
#endregion connect to graph

#region get the users to process
try {
    $users = Get-MGUser -Filter "JobTitle eq 'User to remove caledar items from' and accountEnabled eq true" -All
}
catch {
    Write-Error "Failed to get the users to process"
}
$userCount = ($users | Measure-Object).Count
if ($userCount -lt 1) {
    Write-Error "Less than 1 user to process"
}
if ($userCount -gt $maxUserCount) {
    Write-Error "More than $maxUserCount users returned by filter. User count was $userCount"
}
#endregion get the users to process

#region remove cancelled meetings
$removedCount = 0
foreach ($u in $users) {
    $eventsMG = Get-MgUserEvent -UserId $u.id -Filter $filter -All
    Write-Output "[$($u.userprincipalname)] $($eventsMG.count) events to remove"
    foreach ($e in $eventsMG) {
        if ($e.isCancelled) {
            Write-Output "[$($u.userprincipalname)] removing $($e.Subject) : $($e.start.datetime)"
            Remove-MgUserEvent -UserId $u.Id -EventId $e.Id
            $removedCount ++
        }
        else {
            Write-Warning "[$($u.userprincipalname)] Something odd happened. Event is not cancelled but was returned by the filter. No action taken"
        }
    }
}
#endregion remove cancelled meetings

#region cleanup etc
Write-Output "Processed $($users.count) users"
Write-Output "Removed $removedCount events"
$stopwatch.Stop()
Write-Output "Script completed in $([math]::Round($stopwatch.Elapsed.TotalSeconds,0)) seconds"
$null = Disconnect-MgGraph
$null = Disconnect-AzAccount
#endregion
