# vybe-test: powershell/error_handling/try_finally
$result = "start"
try {
    $result = "try"
} finally {
    $result = $result + "-finally"
}
if ($result -ne "try-finally") {
    Write-Host "FAIL: expected 'try-finally', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
