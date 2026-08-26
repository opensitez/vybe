# vybe-test: powershell/parameters_validate_script/validatescript_empty_guid_throws
function Set-NonNullGuid2 {
    param([ValidateScript({ $_ -ne [guid]::Empty })][guid]$Id)
    return $Id
}
$caught = $false
try {
    $x = Set-NonNullGuid2 -Id ([guid]::Empty)
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Empty GUID should be rejected by ValidateScript"
    exit 1
}
Write-Host "PASS"
exit 0
