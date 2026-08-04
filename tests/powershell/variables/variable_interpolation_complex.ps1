# vybe-test: powershell/variables/variable_interpolation_complex
$obj = [PSCustomObject]@{ Name = "John"; Age = 30 }
$result = "Name: $($obj.Name), Age: $($obj.Age)"
if ($result -ne "Name: John, Age: 30") {
    Write-Host "FAIL: expected 'Name: John, Age: 30', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
