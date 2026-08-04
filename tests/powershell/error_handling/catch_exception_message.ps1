# vybe-test: powershell/error_handling/catch_exception_message
$message = ""
try {
    throw "custom error"
} catch {
    $message = $_.Exception.Message
}
if ($message -ne "custom error") {
    Write-Host "FAIL: expected 'custom error', got '$message'"
    exit 1
}
Write-Host "PASS"
exit 0
