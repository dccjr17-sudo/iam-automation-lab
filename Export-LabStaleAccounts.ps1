#Requires -Version 7.0
[CmdletBinding()]
param([Parameter(Mandatory)] [guid] $ExpectedTenantId, [ValidateRange(1,3650)] [int] $InactiveDays=90, [switch] $UseDeviceCode)
$ErrorActionPreference='Stop'
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
$connected=$false
try {
    $options=@{TenantId=$ExpectedTenantId.ToString();Scopes=@('User.Read.All','AuditLog.Read.All');ContextScope='Process';NoWelcome=$true}
    if ($UseDeviceCode) { $options.UseDeviceCode=$true }
    Connect-MgGraph @options
    $connected=$true
    $context=Get-MgContext
    if (-not $context -or $context.TenantId -ne $ExpectedTenantId.ToString() -or $context.AuthType -ne 'Delegated') { throw 'Expected delegated lab tenant was not established.' }
    foreach ($scope in $options.Scopes) { if ($context.Scopes -notcontains $scope) { throw "Required permission missing: $scope" } }
    Write-Host 'Tenant check: PASSED'
    $now=[datetimeoffset]::UtcNow
    $cutoff=$now.AddDays(-$InactiveDays)
    $next='https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,accountEnabled,createdDateTime,signInActivity&$top=100'
    $seen=[System.Collections.Generic.HashSet[string]]::new()
    $rows=[System.Collections.Generic.List[object]]::new()
    while ($next) {
        $uri=[uri]$next
        if ($uri.Scheme -ne 'https' -or $uri.Host -ne 'graph.microsoft.com' -or $uri.Port -ne 443 -or $uri.UserInfo -or $uri.AbsolutePath -ne '/v1.0/users' -or -not $seen.Add($next)) { throw 'Unexpected user pagination URL.' }
        $page=Invoke-MgGraphRequest -Method GET -Uri $next -OutputType Hashtable
        if (-not $page.ContainsKey('value')) { throw 'Unexpected user response.' }
        foreach ($user in $page.value) {
            $last=$user.signInActivity.lastSuccessfulSignInDateTime
            $state='Unknown history'
            $parsed=[datetimeoffset]::MinValue
            if ($last -and [datetimeoffset]::TryParse([string]$last,[ref]$parsed) -and $parsed.Year -gt 1 -and $parsed -le $now) {
                $state=if ($parsed -lt $cutoff) {'Stale candidate'} else {'Recent successful sign-in'}
            }
            $enabled='Unknown'
            if ($user.accountEnabled -is [bool]) { $enabled=[string]$user.accountEnabled }
            $rows.Add([pscustomobject]@{UserId=$user.id;DisplayName=$user.displayName;UserPrincipalName=$user.userPrincipalName;AccountEnabled=$enabled;CreatedDateTime=$user.createdDateTime;LastSuccessfulSignInDateTime=$last;Assessment=$state;InactiveDays=$InactiveDays;AssessedAtUtc=$now.ToString('o')})
        }
        $next=$page['@odata.nextLink']
    }
    $reports=Join-Path $PSScriptRoot 'reports'
    $null=New-Item -ItemType Directory -Path $reports -Force
    $path=Join-Path $reports ("stale-account-review-$(Get-Date -Format 'yyyyMMdd-HHmmss-fff').csv")
    if ($rows.Count) { $rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding utf8 }
    else { '"UserId","DisplayName","UserPrincipalName","AccountEnabled","CreatedDateTime","LastSuccessfulSignInDateTime","Assessment","InactiveDays","AssessedAtUtc"' | Set-Content -LiteralPath $path -Encoding utf8 }
    Write-Host 'Operation: read-only stale-account review'
    Write-Host "Inactivity threshold: $InactiveDays days"
    Write-Host "Users exported: $($rows.Count)"
    foreach ($state in @('Stale candidate','Recent successful sign-in','Unknown history')) { Write-Host "$($state): $(@($rows | Where-Object Assessment -eq $state).Count)" }
    Write-Host "Saved report: $path"
    Write-Host 'Missing history is unknown, not proof of inactivity. Candidates require review; no accounts were changed.'
} finally { if ($connected) { Disconnect-MgGraph | Out-Null } }
