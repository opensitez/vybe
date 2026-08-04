# vybe-test: powershell/operators/format_operator
$name = "Alice"
$age = 30
$result = "Name: {0}, Age: {1}" -f $name, $age
if ($result -ne "Name: Alice, Age: 30") {
    Write-Host "FAIL: expected 'Name: Alice, Age: 30', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
