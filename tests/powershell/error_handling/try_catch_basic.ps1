# vybe-test: powershell/error_handling/try_catch_basic
$result = "not caught"
try {
    throw "error message"
} catch {
    $result = "caught"
}
if ($result -ne "caught") {
    Write-Host "FAIL: expected 'caught', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
