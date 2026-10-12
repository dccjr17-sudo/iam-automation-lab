#Requires -Version 7.0
# Synthetic responses only; never authenticates or reads live runtime data.
$ErrorActionPreference='Stop'
$tempBase=[IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$testRoot=Join-Path $tempBase ('iam-mover-tests-' + [guid]::NewGuid())
$null=New-Item -ItemType Directory -Path (Join-Path $testRoot 'runtime') -Force
Copy-Item (Join-Path $PSScriptRoot 'Move-LabUser.ps1') $testRoot
$tenant='11111111-1111-1111-1111-111111111111'
$userId='22222222-2222-2222-2222-222222222222'
$teamIds=@{A='33333333-3333-3333-3333-333333333333';B='44444444-4444-4444-4444-444444444444'}
@{TenantId=$tenant;UserId=$userId;UserPrincipalName='lab.auto.joiner-001@synthetic.onmicrosoft.com'} | ConvertTo-Json | Set-Content (Join-Path $testRoot 'runtime/lab-user.local.json')
function Import-Module {}
function Connect-MgGraph { param($TenantId,$Scopes,$ContextScope,$NoWelcome,$UseDeviceCode); $global:moverMock.Scopes=$Scopes; $global:moverMock.Connected=$true }
function Disconnect-MgGraph {}
function Start-Sleep {}
function Get-MgContext { @{TenantId=$(if ($global:moverMock.WrongTenant) {$teamIds.B} else {$tenant});AuthType='Delegated';Scopes=$global:moverMock.Scopes} }
function Invoke-MgGraphRequest {
    param($Method,$Uri,$Body,$ContentType,$OutputType)
    $global:moverMock.Calls.Add("$Method $Uri")
    if ($Uri -match '/users/') { return @{id=$userId;userPrincipalName='lab.auto.joiner-001@synthetic.onmicrosoft.com';department='IAM-Automation-Lab';employeeId='LAB-AUTO-001';accountEnabled=$global:moverMock.Enabled} }
    if ($Uri -match '/groups\?') { return @{value=$(if ($global:moverMock.Collision) {@(@{id=$teamIds.A})} else {@()})} }
    $team=if ($Uri -match [regex]::Escape($teamIds.A)) {'A'} else {'B'}
    if ($Method -eq 'POST' -and $Uri -eq 'https://graph.microsoft.com/v1.0/groups') {
        $team=if ($Body.displayName.EndsWith('-A')) {'A'} else {'B'}
        $global:moverMock.Created[$team]=$true
        return @{id=$teamIds[$team]}
    }
    if ($Uri -match '/members') {
        if ($Method -eq 'GET') {
            return @{value=$(if ($global:moverMock.Member[$team] -or ($team -eq 'A' -and $global:moverMock.StaleSource)) {@(@{id=$userId})} else {@()})}
        }
        if ($Method -eq 'DELETE') {
            if ($Uri -ne "https://graph.microsoft.com/v1.0/groups/$($teamIds[$team])/members/$userId/`$ref") { throw 'Unsafe deletion' }
            $global:moverMock.Member[$team]=$false; return
        }
        if ($Method -eq 'POST') {
            if ($Uri -notlike '*/members/$ref' -or $Body['@odata.id'] -ne "https://graph.microsoft.com/v1.0/directoryObjects/$userId") { throw 'Unsafe addition' }
            if ($global:moverMock.FailAdd) { throw 'Synthetic addition failure' }
            $global:moverMock.Member[$team]=$true; return
        }
    }
    if ($Method -eq 'GET') { return @{id=$teamIds[$team];displayName="LAB-IAM-Automation-Team-$team";description="IAM-Automation-Lab: dedicated mover team $team";securityEnabled=$true;mailEnabled=$false;groupTypes=@();isAssignableToRole=[bool]$global:moverMock.RoleGroup} }
    throw 'Unexpected synthetic request'
}
function Reset-Mock {
    # All mutations below target this uniquely created temporary fixture.
    Get-ChildItem (Join-Path $testRoot 'runtime') -Filter 'mover-*.local.json' | Remove-Item
    $global:moverMock=@{Calls=[System.Collections.Generic.List[string]]::new();Member=@{A=$false;B=$false};Created=@{};Enabled=$false;Connected=$false}
}
function Run-Mover($action,$apply=$true) { & (Join-Path $testRoot 'Move-LabUser.ps1') -ExpectedTenantId $tenant -Action $action -Apply:$apply 6>$null }
function Assert($condition,$message) { if (-not $condition) { throw $message } }
function Expect-Blocked($action) { $blocked=$false; try { Run-Mover $action } catch {$blocked=$true}; Assert $blocked 'Expected refusal' }
$passed=0
function Test-Case($name,[scriptblock]$test) { Reset-Mock; & $test; $script:passed++; Write-Host "PASS: $name" }
try {
    Test-Case 'Preview does not authenticate' { Run-Mover Prepare $false; Assert (-not $global:moverMock.Connected) 'Preview connected' }
    Test-Case 'Enabled user blocks preparation' { $global:moverMock.Enabled=$true; Expect-Blocked Prepare; Assert (@($global:moverMock.Calls | Where-Object {$_ -match '^(POST|DELETE)'}).Count -eq 0) 'Wrote before user validation' }
    Test-Case 'Wrong tenant blocks Graph requests' { $global:moverMock.WrongTenant=$true; Expect-Blocked Prepare; Assert ($global:moverMock.Calls.Count -eq 0) 'Fetched wrong tenant' }
    Test-Case 'Unreferenced group collision blocks adoption' { $global:moverMock.Collision=$true; Expect-Blocked Prepare; Assert ($global:moverMock.Created.Count -eq 0) 'Created despite collision' }
    Test-Case 'Preparation stages A only' { Run-Mover Prepare; Assert ($global:moverMock.Member.A -and -not $global:moverMock.Member.B) 'Wrong preparation' }
    Test-Case 'Move removes A before adding B' { Run-Mover Prepare; $global:moverMock.Calls.Clear(); Run-Mover Move; $writes=@($global:moverMock.Calls | Where-Object {$_ -match '^(POST|DELETE)'}); Assert ($writes.Count -eq 2 -and $writes[0] -match '^DELETE' -and $writes[1] -match '^POST') 'Wrong write order'; Assert (-not $global:moverMock.Member.A -and $global:moverMock.Member.B) 'Wrong moved state' }
    Test-Case 'Repeated move makes no writes' { Run-Mover Prepare; Run-Mover Move; $global:moverMock.Calls.Clear(); Run-Mover Move; Assert (@($global:moverMock.Calls | Where-Object {$_ -match '^(POST|DELETE)'}).Count -eq 0) 'Repeated writes' }
    Test-Case 'Addition failure leaves source removed and can resume' { Run-Mover Prepare; $global:moverMock.FailAdd=$true; Expect-Blocked Move; Assert (-not $global:moverMock.Member.A -and -not $global:moverMock.Member.B) 'Unexpected partial state'; $global:moverMock.FailAdd=$false; Run-Mover Move; Assert $global:moverMock.Member.B 'Resume failed' }
    Test-Case 'Unverified source removal prevents new addition' { Run-Mover Prepare; $global:moverMock.StaleSource=$true; Expect-Blocked Move; Assert (-not $global:moverMock.Member.B) 'Added without removal verification' }
    Test-Case 'Role-assignable group refuses membership changes' { Run-Mover Prepare; $global:moverMock.RoleGroup=$true; $global:moverMock.Calls.Clear(); Expect-Blocked Move; Assert (@($global:moverMock.Calls | Where-Object {$_ -match '^(POST|DELETE)'}).Count -eq 0) 'Changed role group' }
    Test-Case 'Cleanup removes both and prevents restaging by Move' { Run-Mover Prepare; Run-Mover Move; Run-Mover Cleanup; Assert (-not $global:moverMock.Member.A -and -not $global:moverMock.Member.B) 'Cleanup failed'; Expect-Blocked Move }
    Write-Host "Offline mover tests passed: $passed"
} finally {
    $resolved=[IO.Path]::GetFullPath($testRoot)
    if ($resolved.StartsWith($tempBase,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'iam-mover-tests-*') { Remove-Item -LiteralPath $resolved -Recurse -Force }
    Remove-Variable moverMock -Scope Global -ErrorAction SilentlyContinue
}
