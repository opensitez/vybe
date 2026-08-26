# vybe-test: powershell/parameters_alias_attribute/alias_with_guid_parameter
function Set-SessionGuid {
    param([Alias("Token")][guid]$SessionId)
    return $SessionId.ToString()
}
$g = [guid]::NewGuid()
$res = Set-SessionGuid -Token $g
if ($res -ne $g.ToString()) {
    Write-Host "FAIL: Alias on GUID parameter failed"
    exit 1
}
Write-Host "PASS"
exit 0
