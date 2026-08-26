# vybe-test: powershell/parameters_alias_attribute/alias_case_insensitivity
function Set-UserName {
    param(
        [Alias("User")]
        [string]$Username
    )
    return $Username
}
$res = Set-UserName -USER "alice"
if ($res -ne "alice") {
    Write-Host "FAIL: Case-insensitive alias binding failed"
    exit 1
}
Write-Host "PASS"
exit 0
