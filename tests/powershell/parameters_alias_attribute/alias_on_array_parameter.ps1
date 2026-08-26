# vybe-test: powershell/parameters_alias_attribute/alias_on_array_parameter
function Set-ArrayItems {
    param([Alias("Tags")][string[]]$Categories)
    return $Categories.Length
}
$res = Set-ArrayItems -Tags "t1", "t2", "t3"
if ($res -ne 3) {
    Write-Host "FAIL: Alias on array parameter failed"
    exit 1
}
Write-Host "PASS"
exit 0
