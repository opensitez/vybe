# vybe-test: powershell/comparison_and_ordering/basic
if (5 -gt 3 -and 3 -lt 5 -and 4 -ge 4 -and 4 -le 5) {
    # comparisons pass
} else {
    Write-Host 'FAIL: comparison operators returned unexpected result'
    exit 1
}

if ('banana' -gt 'apple' -and 'apple' -lt 'cherry') {
    # ordering pass
} else {
    Write-Host 'FAIL: string ordering failed'
    exit 1
}

Write-Host 'PASS'
exit 0
