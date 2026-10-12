#Requires -Version 7.0
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)] [guid] $ExpectedTenantId,
    [switch] $Apply,
    [switch] $UseDeviceCode
)
$ErrorActionPreference = 'Stop'
$reference = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'runtime/lab-user.local.json') -Raw | ConvertFrom-Json
if ($reference.TenantId -ne $ExpectedTenantId.ToString()) { throw 'Saved lab user belongs to a different tenant.' }
$userId = ([guid] $reference.UserId).ToString()
if ($reference.UserPrincipalName -notmatch '^lab\.auto\.joiner-001@[a-z0-9-]+\.onmicrosoft\.com$') { throw 'Saved user is outside this lab exercise.' }
Write-Host "Planned user: $($reference.UserPrincipalName)"
Write-Host 'Planned change: accountEnabled = False'
Write-Host 'Scope: designated lab user only. Group membership is retained.'
if (-not $Apply) { Write-Host 'Mode: PREVIEW — no connection or changes performed.'; return }
if (-not $PSCmdlet.ShouldProcess($reference.UserPrincipalName, 'Disable designated lab user')) { return }
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
$connected = $false
try {
    $options = @{TenantId=$ExpectedTenantId.ToString();Scopes=@('User.Read.All','User.EnableDisableAccount.All');ContextScope='Process';NoWelcome=$true}
    if ($UseDeviceCode) { $options.UseDeviceCode=$true }
    Connect-MgGraph @options
    $connected = $true
    $context = Get-MgContext
    if (-not $context -or $context.TenantId -ne $ExpectedTenantId.ToString() -or $context.AuthType -ne 'Delegated') { throw 'Expected delegated lab tenant connection was not established.' }
    foreach ($scope in $options.Scopes) { if ($context.Scopes -notcontains $scope) { throw "Required permission missing: $scope" } }
    if ($context.Account -eq $reference.UserPrincipalName) { throw 'Cannot disable the signed-in account.' }
    Write-Host 'Tenant check: PASSED'
    $userUri = "https://graph.microsoft.com/v1.0/users/$userId"
    $readUri = $userUri + '?$select=id,userPrincipalName,department,employeeId,accountEnabled'
    function Assert-LabUser($user) {
        if ($user.id -ne $userId -or $user.userPrincipalName -ne $reference.UserPrincipalName -or
            $user.department -ne 'IAM-Automation-Lab' -or $user.employeeId -ne 'LAB-AUTO-001' -or
            $user.accountEnabled -isnot [bool]) { throw 'User identity, lab markers, or enabled state failed validation.' }
    }
    $user = Invoke-MgGraphRequest -Method GET -Uri $readUri -OutputType Hashtable
    Assert-LabUser $user
    if ($user.accountEnabled -eq $false) {
        Write-Host 'Result: ALREADY DISABLED — no update performed'
    } else {
        $null = Invoke-MgGraphRequest -Method PATCH -Uri $userUri -Body @{accountEnabled=$false} -ContentType 'application/json'
        Write-Host 'Result: DISABLED'
    }
    $verified = $false
    for ($attempt=1; $attempt -le 3; $attempt++) {
        $user = Invoke-MgGraphRequest -Method GET -Uri $readUri -OutputType Hashtable
        Assert-LabUser $user
        if ($user.accountEnabled -eq $false) { $verified=$true; break }
        if ($attempt -lt 3) { Start-Sleep -Seconds 2 }
    }
    if (-not $verified) { throw 'Disable was not confirmed by read-back. Inspect the designated lab user before retrying.' }
    Write-Host 'Disable verification: PASSED'
    Write-Host 'Account enabled: False'
    Write-Host 'This exercise changes the enabled flag only; it does not remove membership or revoke existing sessions.'
} finally {
    if ($connected) { Disconnect-MgGraph | Out-Null }
}
