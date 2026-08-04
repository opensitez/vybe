# vybe-test: powershell/attributes/parameter_alias_attribute
function Test-Func {
    param(
        [Alias("Name")]
        [string]$FullName
    )
    return $FullName
}
$result = Test-Func -Name "Alice"
if ($result -ne "Alice") {
    Write-Host "FAIL: expected Alice, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
