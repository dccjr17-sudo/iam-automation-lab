#Requires -Version 7.0
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)] [guid] $ExpectedTenantId,
    [switch] $Apply,
    [switch] $UseDeviceCode
)

$ErrorActionPreference = 'Stop'
$runtime = Join-Path $PSScriptRoot 'runtime'
$userRef = Get-Content -LiteralPath (Join-Path $runtime 'lab-user.local.json') -Raw | ConvertFrom-Json
if ($userRef.TenantId -ne $ExpectedTenantId.ToString()) { throw 'Saved lab user belongs to a different tenant.' }
$userId = ([guid] $userRef.UserId).ToString()
if ($userRef.UserPrincipalName -notmatch '^lab\.auto\.joiner-001@[a-z0-9-]+\.onmicrosoft\.com$') {
    throw 'Saved sign-in name is outside this lab exercise.'
}
$groupName = 'LAB-IAM-Automation-Joiners'
$marker = 'IAM-Automation-Lab: dedicated joiner membership exercise'
Write-Host "Planned member: $($userRef.UserPrincipalName)"
Write-Host "Planned security group: $groupName"
Write-Host 'Group membership: assigned | Mail enabled: False | Role assignable: False'
Write-Host 'The script assigns no application, directory role, or license to this group.'
if (-not $Apply) { Write-Host 'Mode: PREVIEW — no connection or changes performed.'; return }
if (-not $PSCmdlet.ShouldProcess($groupName, 'Create dedicated lab group if needed and add designated lab user')) { return }

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
$connected = $false
try {
    $options = @{ TenantId=$ExpectedTenantId.ToString(); Scopes=@('User.Read.All','Group.ReadWrite.All'); ContextScope='Process'; NoWelcome=$true }
    if ($UseDeviceCode) { $options.UseDeviceCode=$true }
    Connect-MgGraph @options
    $connected = $true
    $context = Get-MgContext
    if (-not $context -or $context.TenantId -ne $ExpectedTenantId.ToString() -or $context.AuthType -ne 'Delegated') { throw 'Expected delegated lab tenant connection was not established.' }
    foreach ($scope in $options.Scopes) { if ($context.Scopes -notcontains $scope) { throw "Required permission missing: $scope" } }
    Write-Host 'Tenant check: PASSED'
    $user = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/users/$userId" + '?$select=id,userPrincipalName,department,employeeId,accountEnabled') -OutputType Hashtable
    if ($user.userPrincipalName -ne $userRef.UserPrincipalName -or $user.department -ne 'IAM-Automation-Lab' -or $user.employeeId -ne 'LAB-AUTO-001' -or $user.accountEnabled -ne $true) {
        throw 'User markers or enabled state do not match the designated lab user.'
    }
    $groupRefPath = Join-Path $runtime 'lab-group.local.json'
    if (Test-Path -LiteralPath $groupRefPath) {
        $groupRef = Get-Content -LiteralPath $groupRefPath -Raw | ConvertFrom-Json
        if ($groupRef.TenantId -ne $ExpectedTenantId.ToString()) { throw 'Saved group belongs to a different tenant.' }
        $groupId = ([guid] $groupRef.GroupId).ToString()
    } else {
        $lookup = 'https://graph.microsoft.com/v1.0/groups?$filter=' + [uri]::EscapeDataString("displayName eq '$groupName'") + '&$select=id'
        $matches = Invoke-MgGraphRequest -Method GET -Uri $lookup -OutputType Hashtable
        if (-not $matches.ContainsKey('value')) { throw 'Unexpected group lookup response.' }
        if (@($matches.value).Count -gt 0 -or $matches['@odata.nextLink']) { throw 'A group with this name already exists without a local lab reference. Stopped without adopting it.' }
        $group = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/groups' -Body @{
            displayName=$groupName; description=$marker; mailNickname='lab-iam-automation-joiners';
            mailEnabled=$false; securityEnabled=$true; groupTypes=@()
        } -ContentType 'application/json' -OutputType Hashtable
        $groupId = ([guid] $group.id).ToString()
        @{TenantId=$ExpectedTenantId.ToString();GroupId=$groupId;DisplayName=$groupName} | ConvertTo-Json |
            Set-Content -LiteralPath $groupRefPath -Encoding utf8
        Write-Host 'Group: CREATED'
    }
    $groupUri = "https://graph.microsoft.com/v1.0/groups/$groupId"
    $verifiedGroup = Invoke-MgGraphRequest -Method GET -Uri ($groupUri + '?$select=id,displayName,description,securityEnabled,mailEnabled,groupTypes,isAssignableToRole') -OutputType Hashtable
    if ($verifiedGroup.displayName -ne $groupName -or $verifiedGroup.description -ne $marker -or
        $verifiedGroup.securityEnabled -ne $true -or $verifiedGroup.mailEnabled -ne $false -or
        @($verifiedGroup.groupTypes).Count -gt 0 -or $verifiedGroup.isAssignableToRole -eq $true) {
        throw 'Group does not match the dedicated non-role lab security group.'
    }
    function Test-LabMembership {
        $next = $groupUri + '/members?$select=id&$top=100'
        $seen = [System.Collections.Generic.HashSet[string]]::new()
        while ($next) {
            $uri = [uri] $next
            if ($uri.Scheme -ne 'https' -or $uri.Host -ne 'graph.microsoft.com' -or $uri.Port -ne 443 -or
                $uri.UserInfo -or $uri.AbsolutePath -ne "/v1.0/groups/$groupId/members" -or -not $seen.Add($next)) { throw 'Unexpected membership pagination URL.' }
            $page = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType Hashtable
            if (-not $page.ContainsKey('value')) { throw 'Unexpected membership response.' }
            if (@($page.value | Where-Object id -eq $userId).Count -gt 0) { return $true }
            $next = $page['@odata.nextLink']
        }
        return $false
    }
    if (Test-LabMembership) {
        Write-Host 'Membership: ALREADY PRESENT — no duplicate addition'
    } else {
        $null = Invoke-MgGraphRequest -Method POST -Uri ($groupUri + '/members/$ref') -Body @{
            '@odata.id'="https://graph.microsoft.com/v1.0/directoryObjects/$userId"
        } -ContentType 'application/json'
        Write-Host 'Membership: ADDED'
    }
    $present = $false
    for ($attempt=1; $attempt -le 3; $attempt++) {
        if (Test-LabMembership) { $present=$true; break }
        if ($attempt -lt 3) { Start-Sleep -Seconds 2 }
    }
    if (-not $present) { throw 'Membership write completed but read-back did not verify it. Inspect the lab group before retrying.' }
    Write-Host 'Membership verification: PASSED'
    Write-Host "Lab group: $groupName"
    Write-Host 'No licenses, application assignments, or directory roles were assigned by this script.'
} finally {
    if ($connected) { Disconnect-MgGraph | Out-Null }
}
