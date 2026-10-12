#Requires -Version 7.0
[CmdletBinding()]
param([Parameter(Mandatory)] [guid] $ExpectedTenantId, [switch] $UseDeviceCode)
$ErrorActionPreference='Stop'
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
$connected=$false
try {
    $options=@{TenantId=$ExpectedTenantId.ToString();Scopes=@('AuditLog.Read.All');ContextScope='Process';NoWelcome=$true}
    if ($UseDeviceCode) { $options.UseDeviceCode=$true }
    Connect-MgGraph @options
    $connected=$true
    $context=Get-MgContext
    if (-not $context -or $context.TenantId -ne $ExpectedTenantId.ToString() -or $context.AuthType -ne 'Delegated' -or $context.Scopes -notcontains 'AuditLog.Read.All') { throw 'Expected lab tenant and read permission were not established.' }
    Write-Host 'Tenant check: PASSED'
    $next='https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails'
    $seen=[System.Collections.Generic.HashSet[string]]::new()
    $rows=[System.Collections.Generic.List[object]]::new()
    while ($next) {
        $uri=[uri]$next
        if ($uri.Scheme -ne 'https' -or $uri.Host -ne 'graph.microsoft.com' -or $uri.Port -ne 443 -or $uri.UserInfo -or $uri.AbsolutePath -ne '/v1.0/reports/authenticationMethods/userRegistrationDetails' -or -not $seen.Add($next)) { throw 'Unexpected registration pagination URL.' }
        $page=Invoke-MgGraphRequest -Method GET -Uri $next -OutputType Hashtable
        if (-not $page.ContainsKey('value')) { throw 'Unexpected registration response.' }
        foreach ($user in $page.value) {
            $state='Unknown'
            if ($user.isMfaRegistered -is [bool]) { $state=if ($user.isMfaRegistered) {'Registered'} else {'Not registered'} }
            $rows.Add([pscustomobject]@{UserId=$user.id;UserPrincipalName=$user.userPrincipalName;DisplayName=$user.userDisplayName;MfaRegistration=$state;IsMfaCapable=$user.isMfaCapable;IsAdmin=$user.isAdmin;LastUpdatedDateTime=$user.lastUpdatedDateTime})
        }
        $next=$page['@odata.nextLink']
    }
    $reports=Join-Path $PSScriptRoot 'reports'
    $null=New-Item -ItemType Directory -Path $reports -Force
    $path=Join-Path $reports ("mfa-registration-$(Get-Date -Format 'yyyyMMdd-HHmmss-fff').csv")
    if ($rows.Count) { $rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding utf8 }
    else { '"UserId","UserPrincipalName","DisplayName","MfaRegistration","IsMfaCapable","IsAdmin","LastUpdatedDateTime"' | Set-Content -LiteralPath $path -Encoding utf8 }
    Write-Host 'Operation: read-only MFA registration report'
    Write-Host "Users returned: $($rows.Count)"
    Write-Host "Registered: $(@($rows | Where-Object MfaRegistration -eq 'Registered').Count)"
    Write-Host "Not registered: $(@($rows | Where-Object MfaRegistration -eq 'Not registered').Count)"
    Write-Host "Unknown registration: $(@($rows | Where-Object MfaRegistration -eq 'Unknown').Count)"
    Write-Host "Saved report: $path"
    Write-Host 'Registration does not establish MFA enforcement. Disabled users are outside this API coverage; reporting data may lag.'
    Write-Host 'No users or authentication methods changed. Review the local CSV before sharing.'
} finally { if ($connected) { Disconnect-MgGraph | Out-Null } }
