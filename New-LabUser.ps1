#Requires -Version 7.0
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)] [guid] $ExpectedTenantId,
    [Parameter(Mandatory)] [ValidatePattern('^[a-z0-9-]+\.onmicrosoft\.com$')] [string] $LabDomain,
    [switch] $Apply,
    [switch] $UseDeviceCode
)

$ErrorActionPreference = 'Stop'
$upn = "lab.auto.joiner-001@$LabDomain"
Write-Host "Planned user: $upn"
Write-Host 'Display name: LAB Automation Joiner 001'
Write-Host 'Department: IAM-Automation-Lab | Employee marker: LAB-AUTO-001'
Write-Host 'Initial state: enabled | Licenses: none assigned | Roles: none assigned'
if (-not $Apply) {
    Write-Host 'Mode: PREVIEW — no connection or changes performed.'
    return
}
if (-not $PSCmdlet.ShouldProcess($upn, 'Create one designated lab user')) { return }

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
$connected = $false
$body = $null
try {
    $options = @{
        TenantId = $ExpectedTenantId.ToString()
        Scopes = @('User.Create', 'User.Read.All')
        ContextScope = 'Process'
        NoWelcome = $true
    }
    if ($UseDeviceCode) { $options.UseDeviceCode = $true }
    Connect-MgGraph @options
    $connected = $true
    $context = Get-MgContext
    if (-not $context -or $context.TenantId -ne $ExpectedTenantId.ToString() -or $context.AuthType -ne 'Delegated') {
        throw 'Expected delegated lab tenant connection was not established.'
    }
    foreach ($scope in $options.Scopes) {
        if ($context.Scopes -notcontains $scope) { throw "Required permission missing: $scope" }
    }
    Write-Host 'Tenant check: PASSED'
    $lookup = 'https://graph.microsoft.com/v1.0/users?$filter=' +
        [uri]::EscapeDataString("userPrincipalName eq '$upn'") +
        '&$select=id,userPrincipalName,displayName,department,employeeId,accountEnabled'
    $existing = Invoke-MgGraphRequest -Method GET -Uri $lookup -OutputType Hashtable
    if (-not $existing.ContainsKey('value')) { throw 'Unexpected lookup response; stopped before creation.' }
    if (@($existing.value).Count -gt 0) {
        Write-Host 'Creation stopped: this sign-in name already exists. No account changed.'
        return
    }

    # Random password exists only in process memory; it is never printed or saved.
    # This exercise does not sign in as the created user.
    $body = @{
        accountEnabled = $true
        displayName = 'LAB Automation Joiner 001'
        mailNickname = 'lab.auto.joiner-001'
        userPrincipalName = $upn
        department = 'IAM-Automation-Lab'
        employeeId = 'LAB-AUTO-001'
        passwordProfile = @{
            forceChangePasswordNextSignIn = $true
            password = 'Aa1!' + [Convert]::ToBase64String([System.Security.Cryptography.RandomNumberGenerator]::GetBytes(24))
        }
    }
    try {
        $created = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/users' -Body $body -ContentType 'application/json' -OutputType Hashtable
    } catch {
        # Do not emit request bodies or password-containing diagnostic data.
        throw 'User creation did not return success. Inspect the tenant for this exact lab sign-in name before retrying; no password details are logged.'
    } finally {
        if ($body) { $body.passwordProfile.password = $null }
    }
    $newId = [guid] $created.id
    $verifyUri = 'https://graph.microsoft.com/v1.0/users/' + $newId.ToString() +
        '?$select=id,userPrincipalName,displayName,department,employeeId,accountEnabled'
    $verified = $null
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            $verified = Invoke-MgGraphRequest -Method GET -Uri $verifyUri -OutputType Hashtable
            break
        } catch {
            if ($attempt -eq 3) { throw 'Creation returned an ID but read-back was unavailable. Verify the exact lab user in Entra before retrying.' }
            Start-Sleep -Seconds 2
        }
    }
    if ($verified.userPrincipalName -ne $upn -or $verified.department -ne 'IAM-Automation-Lab' -or
        $verified.employeeId -ne 'LAB-AUTO-001' -or $verified.accountEnabled -ne $true) {
        throw 'Created-user verification did not match the intended lab markers. Inspect the user before any further changes.'
    }
    $runtime = Join-Path $PSScriptRoot 'runtime'
    $null = New-Item -ItemType Directory -Path $runtime -Force
    @{
        TenantId = $ExpectedTenantId.ToString()
        UserId = $newId.ToString()
        UserPrincipalName = $upn
        CreatedAtUtc = [datetime]::UtcNow.ToString('o')
    } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $runtime 'lab-user.local.json') -Encoding utf8
    Write-Host 'Result: CREATED and verified'
    Write-Host "Lab user: $upn"
    Write-Host 'Account enabled: True'
    Write-Host 'Local object reference saved. No password, license, or role assignment saved.'
} finally {
    if ($body) { $body.passwordProfile.password = $null }
    if ($connected) { Disconnect-MgGraph | Out-Null }
}
