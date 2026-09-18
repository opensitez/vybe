# vybe-test: powershell/comparison/notcontains_operator
# The -notcontains operator returns true when an array does not contain the specified element
$items = @("apple", "banana", "cherry")

$missingCheck = $items -notcontains "date"
$presentCheck = $items -notcontains "banana"

if (-not $missingCheck) {
    Write-Host "FAIL: expected -notcontains to return true for 'date'"
    exit 1
}

if ($presentCheck) {
    Write-Host "FAIL: expected -notcontains to return false for 'banana'"
    exit 1
}

Write-Host "PASS"
exit 0
