# vybe-test: powershell/numeric_random_number_generation/get_random_from_array_input
$arr = @("apple", "banana", "cherry")
$chosen = $arr | Get-Random
if (-not ($arr -contains $chosen)) {
    Write-Host "FAIL: Get-Random item not in array: $chosen"
    exit 1
}
Write-Host "PASS"
exit 0
