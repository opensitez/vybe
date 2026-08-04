# vybe-test: powershell/error_handling/try_catch_finally
$result = ""
try {
    throw "error"
} catch {
    $result = "caught"
} finally {
    $result = $result + "-finally"
}
if ($result -ne "caught-finally") {
    Write-Host "FAIL: expected 'caught-finally', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
