#Requires -Version 7.0
[CmdletBinding()]
param([Parameter(Mandatory)] [guid] $ExpectedTenantId, [switch] $UseDeviceCode)
$ErrorActionPreference = 'Stop'
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
$connected = $false
try {
    $options = @{TenantId=$ExpectedTenantId.ToString();Scopes=@('RoleManagement.Read.Directory');ContextScope='Process';NoWelcome=$true}
    if ($UseDeviceCode) { $options.UseDeviceCode=$true }
    Connect-MgGraph @options
    $connected=$true
    $context=Get-MgContext
    if (-not $context -or $context.TenantId -ne $ExpectedTenantId.ToString() -or $context.AuthType -ne 'Delegated' -or $context.Scopes -notcontains 'RoleManagement.Read.Directory') { throw 'Expected lab tenant and read permission were not established.' }
    Write-Host 'Tenant check: PASSED'
    $next='https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?$expand=principal'
    $roleCache=@{}
    $seen=[System.Collections.Generic.HashSet[string]]::new()
    $rows=[System.Collections.Generic.List[object]]::new()
    while ($next) {
        $uri=[uri]$next
        if ($uri.Scheme -ne 'https' -or $uri.Host -ne 'graph.microsoft.com' -or $uri.Port -ne 443 -or $uri.UserInfo -or $uri.AbsolutePath -ne '/v1.0/roleManagement/directory/roleAssignments' -or -not $seen.Add($next)) { throw 'Unexpected role assignment pagination URL.' }
        $page=Invoke-MgGraphRequest -Method GET -Uri $next -OutputType Hashtable
        if (-not $page.ContainsKey('value')) { throw 'Unexpected role assignment response.' }
        foreach ($assignment in $page.value) {
            $principal=$assignment.principal
            $roleId=([guid]$assignment.roleDefinitionId).ToString()
            if (-not $roleCache.ContainsKey($roleId)) {
                $roleCache[$roleId]=Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions/$roleId" + '?$select=id,displayName') -OutputType Hashtable
            }
            $role=$roleCache[$roleId]
            if (-not $principal) { $principal=@{} }
            if (-not $role) { $role=@{} }
            $rows.Add([pscustomobject]@{
                AssignmentId=$assignment.id;PrincipalId=$assignment.principalId;
                PrincipalType=$principal['@odata.type'];PrincipalDisplayName=$principal.displayName;
                UserPrincipalName=$principal.userPrincipalName;RoleDefinitionId=$assignment.roleDefinitionId;
                RoleName=$role.displayName;DirectoryScopeId=$assignment.directoryScopeId;AppScopeId=$assignment.appScopeId
            })
        }
        $next=$page['@odata.nextLink']
    }
    $reports=Join-Path $PSScriptRoot 'reports'
    $null=New-Item -ItemType Directory -Path $reports -Force
    $path=Join-Path $reports ("active-directory-role-assignments-$(Get-Date -Format 'yyyyMMdd-HHmmss-fff').csv")
    if ($rows.Count) { $rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding utf8 }
    else { '"AssignmentId","PrincipalId","PrincipalType","PrincipalDisplayName","UserPrincipalName","RoleDefinitionId","RoleName","DirectoryScopeId","AppScopeId"' | Set-Content -LiteralPath $path -Encoding utf8 }
    Write-Host 'Operation: read-only active directory role assignment report'
    Write-Host "Assignments exported: $($rows.Count)"
    Write-Host "Saved report: $path"
    Write-Host 'Coverage: direct active role assignments. Group members and PIM eligible assignments are not expanded.'
    Write-Host 'No roles, users, or groups changed. Review the local CSV before sharing.'
} finally { if ($connected) { Disconnect-MgGraph | Out-Null } }
