# vybe-test: powershell/parameters_validate_not_null_or_empty/validatenotnullorempty_on_guid_parameter
function Set-GuidVal {
    param([ValidateNotNullOrEmpty()][guid]$Id)
    return $Id.ToString()
}
$g = [guid]::NewGuid()
$res = Set-GuidVal -Id $g
if ($res -ne $g.ToString()) {
    Write-Host "FAIL: Guid parameter ValidateNotNullOrEmpty failed"
    exit 1
}
Write-Host "PASS"
exit 0
