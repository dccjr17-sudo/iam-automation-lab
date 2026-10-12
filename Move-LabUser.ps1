#Requires -Version 7.0
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)] [guid] $ExpectedTenantId,
    [ValidateSet('Prepare','Move','Cleanup')] [string] $Action='Move',
    [switch] $Apply,
    [switch] $UseDeviceCode
)
$ErrorActionPreference='Stop'
$runtime=Join-Path $PSScriptRoot 'runtime'
$userRef=Get-Content -LiteralPath (Join-Path $runtime 'lab-user.local.json') -Raw | ConvertFrom-Json
if ($userRef.TenantId -ne $ExpectedTenantId.ToString() -or $userRef.UserPrincipalName -notmatch '^lab\.auto\.joiner-001@[a-z0-9-]+\.onmicrosoft\.com$') { throw 'Saved user reference is outside the designated tenant or lab.' }
$userId=([guid]$userRef.UserId).ToString()
$names=@{A='LAB-IAM-Automation-Team-A';B='LAB-IAM-Automation-Team-B'}
Write-Host "Planned user: $($userRef.UserPrincipalName)"
Write-Host "Action: $Action | Source: $($names.A) | Destination: $($names.B)"
Write-Host 'User must remain disabled. No applications, licenses, or roles are assigned by this script.'
switch ($Action) {
    Prepare { Write-Host 'Plan: create two dedicated lab groups if needed and stage membership in Team A.' }
    Move { Write-Host 'Plan: remove and verify Team A membership, then add and verify Team B membership.' }
    Cleanup { Write-Host 'Plan: remove membership from both mover groups; retain user and groups.' }
}
if (-not $Apply) { Write-Host 'Mode: PREVIEW — no connection or changes performed.'; return }
if (-not $PSCmdlet.ShouldProcess($userRef.UserPrincipalName,"$Action designated mover memberships")) { return }
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
$connected=$false
try {
    $scopes=if ($Action -eq 'Prepare') {@('User.Read.All','Group.ReadWrite.All')} else {@('User.Read.All','Group.Read.All','GroupMember.ReadWrite.All')}
    $options=@{TenantId=$ExpectedTenantId.ToString();Scopes=$scopes;ContextScope='Process';NoWelcome=$true}
    if ($UseDeviceCode) { $options.UseDeviceCode=$true }
    Connect-MgGraph @options
    $connected=$true
    $context=Get-MgContext
    if (-not $context -or $context.TenantId -ne $ExpectedTenantId.ToString() -or $context.AuthType -ne 'Delegated') { throw 'Expected delegated lab tenant was not established.' }
    foreach ($scope in $scopes) { if ($context.Scopes -notcontains $scope) { throw "Required permission missing: $scope" } }
    Write-Host 'Tenant check: PASSED'
    $userUri="https://graph.microsoft.com/v1.0/users/$userId" + '?$select=id,userPrincipalName,department,employeeId,accountEnabled'
    function Assert-LabUser {
        $user=Invoke-MgGraphRequest -Method GET -Uri $userUri -OutputType Hashtable
        if ($user.id -ne $userId -or $user.userPrincipalName -ne $userRef.UserPrincipalName -or $user.department -ne 'IAM-Automation-Lab' -or $user.employeeId -ne 'LAB-AUTO-001' -or $user.accountEnabled -isnot [bool] -or $user.accountEnabled -ne $false) { throw 'Expected disabled lab user and markers were not verified.' }
    }
    Assert-LabUser
    $ids=@{}
    foreach ($team in @('A','B')) {
        $path=Join-Path $runtime "mover-team-$team.local.json"
        $marker="IAM-Automation-Lab: dedicated mover team $team"
        if (Test-Path -LiteralPath $path) {
            $reference=Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
            if ($reference.TenantId -ne $ExpectedTenantId.ToString() -or $reference.DisplayName -ne $names[$team]) { throw 'Mover group reference does not match the lab.' }
            $ids[$team]=([guid]$reference.GroupId).ToString()
        } else {
            if ($Action -ne 'Prepare') { throw 'Mover group reference missing. Complete Prepare first.' }
            $lookup='https://graph.microsoft.com/v1.0/groups?$filter=' + [uri]::EscapeDataString("displayName eq '$($names[$team])'") + '&$select=id'
            $existing=Invoke-MgGraphRequest -Method GET -Uri $lookup -OutputType Hashtable
            if (-not $existing.ContainsKey('value') -or @($existing.value).Count -gt 0 -or $existing['@odata.nextLink']) { throw 'Unexpected lookup or existing unreferenced mover group. Inspect before retrying; no adoption performed.' }
            $created=Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/groups' -Body @{displayName=$names[$team];description=$marker;mailNickname="lab-iam-automation-team-$team";mailEnabled=$false;securityEnabled=$true;groupTypes=@()} -ContentType 'application/json' -OutputType Hashtable
            $ids[$team]=([guid]$created.id).ToString()
            @{TenantId=$ExpectedTenantId.ToString();GroupId=$ids[$team];DisplayName=$names[$team]} | ConvertTo-Json | Set-Content -LiteralPath $path -Encoding utf8
            Write-Host "Team $team group: CREATED"
        }
        $group=Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/groups/$($ids[$team])" + '?$select=id,displayName,description,securityEnabled,mailEnabled,groupTypes,isAssignableToRole') -OutputType Hashtable
        if ($group.id -ne $ids[$team] -or $group.displayName -ne $names[$team] -or $group.description -ne $marker -or $group.securityEnabled -ne $true -or $group.mailEnabled -ne $false -or @($group.groupTypes).Count -gt 0 -or $group.isAssignableToRole -eq $true) { throw 'Mover group properties failed validation.' }
    }
    if ($ids.A -eq $ids.B) { throw 'Source and destination groups must differ.' }
    function Test-Membership([string]$team) {
        $next="https://graph.microsoft.com/v1.0/groups/$($ids[$team])/members" + '?$select=id&$top=100'
        $seen=[System.Collections.Generic.HashSet[string]]::new()
        while ($next) {
            $uri=[uri]$next
            if ($uri.Scheme -ne 'https' -or $uri.Host -ne 'graph.microsoft.com' -or $uri.Port -ne 443 -or $uri.UserInfo -or $uri.AbsolutePath -ne "/v1.0/groups/$($ids[$team])/members" -or -not $seen.Add($next)) { throw 'Unexpected membership pagination URL.' }
            $page=Invoke-MgGraphRequest -Method GET -Uri $next -OutputType Hashtable
            if (-not $page.ContainsKey('value')) { throw 'Unexpected membership response.' }
            if (@($page.value | Where-Object id -eq $userId).Count) { return $true }
            $next=$page['@odata.nextLink']
        }
        return $false
    }
    function Assert-Membership([string]$team,[bool]$expected) {
        for ($attempt=1; $attempt -le 3; $attempt++) {
            if ((Test-Membership $team) -eq $expected) { return }
            if ($attempt -lt 3) { Start-Sleep -Seconds 2 }
        }
        throw "Team $team membership did not verify. Inspect both groups before retrying."
    }
    function Add-Membership([string]$team) {
        if (-not (Test-Membership $team)) {
            $null=Invoke-MgGraphRequest -Method POST -Uri ("https://graph.microsoft.com/v1.0/groups/$($ids[$team])/members/" + '$ref') -Body @{'@odata.id'="https://graph.microsoft.com/v1.0/directoryObjects/$userId"} -ContentType 'application/json'
            Write-Host "Team $team membership: ADDED"
        } else { Write-Host "Team $team membership: ALREADY PRESENT" }
        Assert-Membership $team $true
    }
    function Remove-Membership([string]$team) {
        if (Test-Membership $team) {
            $null=Invoke-MgGraphRequest -Method DELETE -Uri ("https://graph.microsoft.com/v1.0/groups/$($ids[$team])/members/$userId/" + '$ref')
            Write-Host "Team $team membership: REMOVED"
        } else { Write-Host "Team $team membership: ALREADY ABSENT" }
        Assert-Membership $team $false
    }
    $statePath=Join-Path $runtime 'mover-state.local.json'
    if ($Action -eq 'Prepare') {
        if (Test-Path -LiteralPath $statePath) { throw 'Preparation already recorded. Use Move or Cleanup; preparation does not reset an exercise.' }
        if (Test-Membership 'B') { throw 'Destination already contains user. Inspect the exercise state before preparing.' }
        Add-Membership 'A'
        @{TenantId=$ExpectedTenantId.ToString();UserId=$userId;SourceId=$ids.A;DestinationId=$ids.B;PreparedAtUtc=[datetimeoffset]::UtcNow.ToString('o')} | ConvertTo-Json | Set-Content -LiteralPath $statePath -Encoding utf8
        Assert-Membership 'B' $false
        Write-Host 'Preparation verification: PASSED'
    } elseif ($Action -eq 'Move') {
        $state=Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
        if ($state.TenantId -ne $ExpectedTenantId.ToString() -or $state.UserId -ne $userId -or $state.SourceId -ne $ids.A -or $state.DestinationId -ne $ids.B -or $state.CleanedAtUtc) { throw 'Prepared mover state does not match or has been cleaned up.' }
        Remove-Membership 'A'
        Assert-LabUser
        Add-Membership 'B'
        Assert-Membership 'A' $false
        Write-Host 'Mover verification: PASSED | Team A: ABSENT | Team B: PRESENT'
    } else {
        Remove-Membership 'A'
        Remove-Membership 'B'
        if (Test-Path -LiteralPath $statePath) {
            $state=Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
            $state | Add-Member -NotePropertyName CleanedAtUtc -NotePropertyValue ([datetimeoffset]::UtcNow.ToString('o')) -Force
            $state | ConvertTo-Json | Set-Content -LiteralPath $statePath -Encoding utf8
        }
        Write-Host 'Cleanup verification: PASSED | Team A: ABSENT | Team B: ABSENT'
    }
    Assert-LabUser
    Write-Host 'User retained and disabled: VERIFIED'
    Write-Host 'Membership changes are separate Graph requests, not a transaction. Inspect partial results after errors; rerun Move to reconcile before cleanup.'
} finally { if ($connected) { Disconnect-MgGraph | Out-Null } }
