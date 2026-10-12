#Requires -Version 7.0
[CmdletBinding(SupportsShouldProcess)]
param([Parameter(Mandatory)] [guid] $ExpectedTenantId, [switch] $Apply, [switch] $UseDeviceCode)
$ErrorActionPreference='Stop'
$runtime=Join-Path $PSScriptRoot 'runtime'
$userRef=Get-Content -LiteralPath (Join-Path $runtime 'lab-user.local.json') -Raw | ConvertFrom-Json
$groupRef=Get-Content -LiteralPath (Join-Path $runtime 'lab-group.local.json') -Raw | ConvertFrom-Json
if ($userRef.TenantId -ne $ExpectedTenantId.ToString() -or $groupRef.TenantId -ne $ExpectedTenantId.ToString()) { throw 'Saved references belong to a different tenant.' }
$userId=([guid]$userRef.UserId).ToString()
$groupId=([guid]$groupRef.GroupId).ToString()
if ($userRef.UserPrincipalName -notmatch '^lab\.auto\.joiner-001@[a-z0-9-]+\.onmicrosoft\.com$' -or $groupRef.DisplayName -ne 'LAB-IAM-Automation-Joiners') { throw 'References are outside the designated lab exercise.' }
Write-Host "Planned user: $($userRef.UserPrincipalName)"
Write-Host 'Planned change: remove membership from LAB-IAM-Automation-Joiners'
Write-Host 'Prerequisite: lab user must already be disabled. User and group are retained.'
if (-not $Apply) { Write-Host 'Mode: PREVIEW — no connection or changes performed.'; return }
if (-not $PSCmdlet.ShouldProcess($userRef.UserPrincipalName,'Remove designated lab group membership')) { return }
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
$connected=$false
try {
    $options=@{TenantId=$ExpectedTenantId.ToString();Scopes=@('User.Read.All','GroupMember.ReadWrite.All','Group.Read.All');ContextScope='Process';NoWelcome=$true}
    if ($UseDeviceCode) { $options.UseDeviceCode=$true }
    Connect-MgGraph @options
    $connected=$true
    $context=Get-MgContext
    if (-not $context -or $context.TenantId -ne $ExpectedTenantId.ToString() -or $context.AuthType -ne 'Delegated') { throw 'Expected delegated tenant was not established.' }
    foreach ($scope in $options.Scopes) { if ($context.Scopes -notcontains $scope) { throw "Required permission missing: $scope" } }
    Write-Host 'Tenant check: PASSED'
    $userUri="https://graph.microsoft.com/v1.0/users/$userId" + '?$select=id,userPrincipalName,department,employeeId,accountEnabled'
    function Assert-DisabledLabUser($user) {
        if ($user.id -ne $userId -or $user.userPrincipalName -ne $userRef.UserPrincipalName -or $user.department -ne 'IAM-Automation-Lab' -or $user.employeeId -ne 'LAB-AUTO-001' -or $user.accountEnabled -isnot [bool] -or $user.accountEnabled -ne $false) { throw 'Expected disabled lab user was not verified.' }
    }
    Assert-DisabledLabUser (Invoke-MgGraphRequest -Method GET -Uri $userUri -OutputType Hashtable)
    $groupUri="https://graph.microsoft.com/v1.0/groups/$groupId"
    $group=Invoke-MgGraphRequest -Method GET -Uri ($groupUri + '?$select=id,displayName,description,securityEnabled,mailEnabled,groupTypes,isAssignableToRole') -OutputType Hashtable
    if ($group.id -ne $groupId -or $group.displayName -ne 'LAB-IAM-Automation-Joiners' -or $group.description -ne 'IAM-Automation-Lab: dedicated joiner membership exercise' -or $group.securityEnabled -ne $true -or $group.mailEnabled -ne $false -or @($group.groupTypes).Count -gt 0 -or $group.isAssignableToRole -eq $true) { throw 'Designated non-role lab group was not verified.' }
    function Test-LabMembership {
        $next=$groupUri + '/members?$select=id&$top=100'
        $seen=[System.Collections.Generic.HashSet[string]]::new()
        while ($next) {
            $uri=[uri]$next
            if ($uri.Scheme -ne 'https' -or $uri.Host -ne 'graph.microsoft.com' -or $uri.Port -ne 443 -or $uri.UserInfo -or $uri.AbsolutePath -ne "/v1.0/groups/$groupId/members" -or -not $seen.Add($next)) { throw 'Unexpected membership pagination URL.' }
            $page=Invoke-MgGraphRequest -Method GET -Uri $next -OutputType Hashtable
            if (-not $page.ContainsKey('value')) { throw 'Unexpected membership response.' }
            if (@($page.value | Where-Object id -eq $userId).Count) { return $true }
            $next=$page['@odata.nextLink']
        }
        return $false
    }
    if (Test-LabMembership) {
        $removeUri=$groupUri + "/members/$userId/" + '$ref'
        $null=Invoke-MgGraphRequest -Method DELETE -Uri $removeUri
        Write-Host 'Membership: REMOVED'
    } else { Write-Host 'Membership: ALREADY ABSENT — no removal performed' }
    $absent=$false
    for ($attempt=1; $attempt -le 3; $attempt++) {
        if (-not (Test-LabMembership)) { $absent=$true; break }
        if ($attempt -lt 3) { Start-Sleep -Seconds 2 }
    }
    if (-not $absent) { throw 'Membership removal was not verified. Inspect the lab group.' }
    Assert-DisabledLabUser (Invoke-MgGraphRequest -Method GET -Uri $userUri -OutputType Hashtable)
    Write-Host 'Removal verification: PASSED'
    Write-Host 'User retained and disabled: VERIFIED'
    Write-Host 'Only designated group membership removed. Existing sessions and other access are outside this exercise.'
} finally { if ($connected) { Disconnect-MgGraph | Out-Null } }
