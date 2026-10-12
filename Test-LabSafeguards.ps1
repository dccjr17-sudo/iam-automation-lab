#Requires -Version 7.0
# Offline synthetic Graph responses only. No Graph module or tenant connection.
$ErrorActionPreference='Stop'
$testRoot=Join-Path ([IO.Path]::GetTempPath()) ('iam-lab-tests-' + [guid]::NewGuid())
$null=New-Item -ItemType Directory -Path (Join-Path $testRoot 'runtime') -Force
Get-ChildItem $PSScriptRoot -Filter '*.ps1' | Where-Object Name -ne $MyInvocation.MyCommand.Name | Copy-Item -Destination $testRoot
$tenant='11111111-1111-1111-1111-111111111111'
$userId='22222222-2222-2222-2222-222222222222'
$groupId='33333333-3333-3333-3333-333333333333'
@{TenantId=$tenant;UserId=$userId;UserPrincipalName='lab.auto.joiner-001@synthetic.onmicrosoft.com'} | ConvertTo-Json | Set-Content (Join-Path $testRoot 'runtime/lab-user.local.json')
@{TenantId=$tenant;GroupId=$groupId;DisplayName='LAB-IAM-Automation-Joiners'} | ConvertTo-Json | Set-Content (Join-Path $testRoot 'runtime/lab-group.local.json')
function Import-Module { param($Name) }
function Connect-MgGraph { param($TenantId,$Scopes,$ContextScope,$NoWelcome,$UseDeviceCode); $global:iamMock.Scopes=$Scopes; $global:iamMock.Connected=$true }
function Disconnect-MgGraph { $global:iamMock.Disconnected=$true }
function Get-MgContext {
    @{TenantId=$(if ($global:iamMock.WrongTenant) {'44444444-4444-4444-4444-444444444444'} else {$tenant});AuthType='Delegated';Scopes=$(if ($global:iamMock.MissingScope) {@()} else {$global:iamMock.Scopes});Account='admin@synthetic.onmicrosoft.com'}
}
function Invoke-MgGraphRequest {
    param($Method,$Uri,$Body,$ContentType,$OutputType)
    $global:iamMock.Calls.Add("$Method $Uri")
    if ($Method -eq 'PATCH') { $global:iamMock.Enabled=$false; return }
    if ($Method -eq 'DELETE') {
        if ($Uri -ne "https://graph.microsoft.com/v1.0/groups/$groupId/members/$userId/`$ref") { throw 'Unsafe deletion URL' }
        $global:iamMock.Member=$false; return
    }
    if ($Method -ne 'GET') { throw 'Unexpected write' }
    if ($Uri -match '/members\?') { return @{value=$(if ($global:iamMock.Member) {@(@{id=$userId})} else {@()})} }
    if ($Uri -match '/groups/') { return @{id=$groupId;displayName='LAB-IAM-Automation-Joiners';description=$(if ($global:iamMock.WrongGroup) {'unrelated'} else {'IAM-Automation-Lab: dedicated joiner membership exercise'});securityEnabled=$true;mailEnabled=$false;groupTypes=@();isAssignableToRole=$false} }
    if ($Uri -match "/users/$userId") { return @{id=$userId;userPrincipalName='lab.auto.joiner-001@synthetic.onmicrosoft.com';department=$(if ($global:iamMock.WrongMarker) {'unrelated'} else {'IAM-Automation-Lab'});employeeId='LAB-AUTO-001';accountEnabled=$global:iamMock.Enabled} }
    if ($Uri -match '/users\?') {
        $page=@{value=@(@{id='a';accountEnabled=$null;signInActivity=@{}},@{id='b';accountEnabled=$true;signInActivity=@{lastSuccessfulSignInDateTime=[datetimeoffset]::UtcNow.AddDays(-120).ToString('o')}},@{id='c';accountEnabled=$false;signInActivity=@{lastSuccessfulSignInDateTime=[datetimeoffset]::UtcNow.AddDays(-1).ToString('o')}})}
        if ($global:iamMock.BadPage) { $page['@odata.nextLink']='https://example.test/v1.0/users' }
        return $page
    }
    if ($Uri -match 'userRegistrationDetails') { return @{value=@(@{id='a';isMfaRegistered=$null},@{id='b';isMfaRegistered=$true},@{id='c';isMfaRegistered=$false})} }
    throw "Unexpected synthetic request: $Uri"
}
function Reset-Mock {
    $global:iamMock=@{Enabled=$true;Member=$true;Calls=[System.Collections.Generic.List[string]]::new();Connected=$false;Disconnected=$false}
}
function Assert($condition,$message) { if (-not $condition) { throw $message } }
function Run-Script($name,$apply=$false) {
    $args=@{ExpectedTenantId=$tenant}
    if ($apply) { $args.Apply=$true }
    & (Join-Path $testRoot $name) @args 6>$null
}
function Expect-Blocked($name) {
    $blocked=$false
    try { Run-Script $name $true } catch { $blocked=$true }
    Assert $blocked 'Expected refusal did not occur'
    Assert (@($global:iamMock.Calls | Where-Object {$_ -match '^(PATCH|DELETE|POST) '}).Count -eq 0) 'Refusal performed a write'
}
$passed=0
function Test-Case($name,[scriptblock]$test) { Reset-Mock; & $test; Write-Host "PASS: $name"; $script:passed++ }
try {
    Test-Case 'Disable preview does not connect' { Run-Script 'Disable-LabUser.ps1'; Assert (-not $global:iamMock.Connected) 'Preview connected' }
    Test-Case 'Wrong tenant refuses writes' { $global:iamMock.WrongTenant=$true; Expect-Blocked 'Disable-LabUser.ps1' }
    Test-Case 'Missing permission refuses writes' { $global:iamMock.MissingScope=$true; Expect-Blocked 'Disable-LabUser.ps1' }
    Test-Case 'Wrong user marker refuses writes' { $global:iamMock.WrongMarker=$true; Expect-Blocked 'Disable-LabUser.ps1' }
    Test-Case 'Disable writes once and verifies' { Run-Script 'Disable-LabUser.ps1' $true; Assert (-not $global:iamMock.Enabled) 'Not disabled'; Assert (@($global:iamMock.Calls | Where-Object {$_ -match '^PATCH '}).Count -eq 1) 'Unexpected patch count'; Assert $global:iamMock.Disconnected 'Not disconnected' }
    Test-Case 'Already disabled skips update' { $global:iamMock.Enabled=$false; Run-Script 'Disable-LabUser.ps1' $true; Assert (@($global:iamMock.Calls | Where-Object {$_ -match '^PATCH '}).Count -eq 0) 'Repeated patch' }
    Test-Case 'Leaver cleanup rejects enabled user' { Expect-Blocked 'Remove-LabGroupMembership.ps1' }
    Test-Case 'Leaver cleanup rejects unrelated group' { $global:iamMock.Enabled=$false; $global:iamMock.WrongGroup=$true; Expect-Blocked 'Remove-LabGroupMembership.ps1' }
    Test-Case 'Removal uses reference URL and retains user' { $global:iamMock.Enabled=$false; Run-Script 'Remove-LabGroupMembership.ps1' $true; Assert (-not $global:iamMock.Member) 'Still member'; Assert (@($global:iamMock.Calls | Where-Object {$_ -match '^DELETE '}).Count -eq 1) 'Unexpected deletion count' }
    Test-Case 'Already absent skips deletion' { $global:iamMock.Enabled=$false; $global:iamMock.Member=$false; Run-Script 'Remove-LabGroupMembership.ps1' $true; Assert (@($global:iamMock.Calls | Where-Object {$_ -match '^DELETE '}).Count -eq 0) 'Repeated deletion' }
    Test-Case 'Stale assessment separates missing, old, recent history' { Run-Script 'Export-LabStaleAccounts.ps1'; $csv=Get-ChildItem (Join-Path $testRoot 'reports/stale*.csv') | Sort-Object LastWriteTime | Select-Object -Last 1; $rows=@(Import-Csv $csv.FullName); Assert ($rows[0].Assessment -eq 'Unknown history' -and $rows[1].Assessment -eq 'Stale candidate' -and $rows[2].Assessment -eq 'Recent successful sign-in') 'Wrong assessments' }
    Test-Case 'Foreign pagination is rejected before following it' { $global:iamMock.BadPage=$true; $blocked=$false; try { Run-Script 'Export-LabUsers.ps1' } catch { $blocked=$true }; Assert $blocked 'Bad page accepted'; Assert ($global:iamMock.Calls.Count -eq 1) 'Foreign page fetched' }
    Test-Case 'MFA null stays unknown' { Run-Script 'Export-LabMfaRegistration.ps1'; $csv=Get-ChildItem (Join-Path $testRoot 'reports/mfa*.csv') | Sort-Object LastWriteTime | Select-Object -Last 1; $rows=@(Import-Csv $csv.FullName); Assert ($rows[0].MfaRegistration -eq 'Unknown' -and $rows[1].MfaRegistration -eq 'Registered' -and $rows[2].MfaRegistration -eq 'Not registered') 'Wrong registration assessment' }
    Write-Host "Offline safeguard tests passed: $passed"
} finally {
    # Delete only this freshly created, uniquely named temporary test directory.
    $resolved=[IO.Path]::GetFullPath($testRoot)
    $temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    if ($resolved.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'iam-lab-tests-*') { Remove-Item -LiteralPath $resolved -Recurse -Force }
    Remove-Variable iamMock -Scope Global -ErrorAction SilentlyContinue
}
