#Requires -Version 7.0
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [guid] $ExpectedTenantId,
    [switch] $UseDeviceCode
)

$ErrorActionPreference = 'Stop'
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
$connected = $false

try {
    $connectOptions = @{
        TenantId = $ExpectedTenantId.ToString()
        Scopes = @('User.Read.All')
        ContextScope = 'Process'
        NoWelcome = $true
    }
    if ($UseDeviceCode) { $connectOptions.UseDeviceCode = $true }
    Connect-MgGraph @connectOptions
    $connected = $true
    $context = Get-MgContext
    if (-not $context -or $context.TenantId -ne $ExpectedTenantId.ToString()) {
        throw 'Connected tenant does not match the expected lab tenant. No report was requested.'
    }
    if ($context.AuthType -ne 'Delegated' -or $context.Scopes -notcontains 'User.Read.All') {
        throw 'The report requires delegated User.Read.All permission.'
    }
    Write-Host 'Tenant check: PASSED'
    Write-Host 'Operation: read-only user inventory'

    $users = [System.Collections.Generic.List[object]]::new()
    $seenPages = [System.Collections.Generic.HashSet[string]]::new()
    $next = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,accountEnabled,userType,department,createdDateTime&$top=100'
    while ($next) {
        $pageUri = [uri] $next
        if ($pageUri.Scheme -ne 'https' -or $pageUri.Host -ne 'graph.microsoft.com' -or
            $pageUri.AbsolutePath -ne '/v1.0/users' -or $pageUri.Port -ne 443 -or
            $pageUri.UserInfo -or -not $seenPages.Add($next)) {
            throw 'Unexpected or repeated pagination URL. Report stopped.'
        }
        $page = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType Hashtable
        if (-not $page.ContainsKey('value')) { throw 'User response did not contain a collection.' }
        foreach ($user in $page.value) {
            $enabled = if ($null -eq $user.accountEnabled) { 'Unknown' } else { [string] [bool] $user.accountEnabled }
            $users.Add([pscustomobject]@{
                Id = $user.id
                DisplayName = $user.displayName
                UserPrincipalName = $user.userPrincipalName
                AccountEnabled = $enabled
                UserType = $user.userType
                Department = $user.department
                CreatedDateTime = $user.createdDateTime
            })
        }
        $next = $page['@odata.nextLink']
    }

    $reports = Join-Path $PSScriptRoot 'reports'
    $null = New-Item -ItemType Directory -Path $reports -Force
    $stamp = [datetime]::UtcNow.ToString('yyyyMMdd-HHmmss-fff')
    $reportPath = Join-Path $reports "user-inventory-$stamp.csv"
    if ($users.Count -gt 0) {
        $users | Sort-Object UserPrincipalName | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding utf8
    } else {
        '"Id","DisplayName","UserPrincipalName","AccountEnabled","UserType","Department","CreatedDateTime"' |
            Set-Content -LiteralPath $reportPath -Encoding utf8
    }
    Write-Host "Users exported: $($users.Count)"
    Write-Host "Enabled accounts: $(@($users | Where-Object AccountEnabled -eq 'True').Count)"
    Write-Host "Disabled accounts: $(@($users | Where-Object AccountEnabled -eq 'False').Count)"
    Write-Host "Unknown enabled state: $(@($users | Where-Object AccountEnabled -eq 'Unknown').Count)"
    Write-Host "Saved report: $reportPath"
    Write-Host 'No users or groups were changed. Review the local CSV before sharing it.'
} finally {
    if ($connected) { Disconnect-MgGraph | Out-Null }
}
