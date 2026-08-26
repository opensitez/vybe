# vybe-test: powershell/parameters_validate_script/validatescript_guid_not_empty
function Set-NonNullGuid {
    param([ValidateScript({ $_ -ne [guid]::Empty })][guid]$Id)
    return $Id.ToString()
}
$g = [guid]::NewGuid()
$res = Set-NonNullGuid -Id $g
if ($res -ne $g.ToString()) {
    Write-Host "FAIL: ValidateScript non-empty GUID failed"
    exit 1
}
Write-Host "PASS"
exit 0
