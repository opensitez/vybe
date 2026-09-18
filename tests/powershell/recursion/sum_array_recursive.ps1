# vybe-test: powershell/recursion/sum_array_recursive
function Get-Sum {
    param([array]$arr, [int]$index = 0)
    if ($index -ge $arr.Count) {
        return 0
    }
    return $arr[$index] + (Get-Sum $arr ($index + 1))
}

$result = Get-Sum @(1, 2, 3, 4, 5)
if ($result -ne 15) {
    Write-Host "FAIL: expected 15, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
