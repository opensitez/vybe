# vybe-test: powershell/parameters_validate_pattern/validatepattern_error_message_contains_pattern
function Set-Zip {
    param([ValidatePattern('^\d{5}$')][string]$Zip)
    return $Zip
}
$msg = ""
try {
    $x = Set-Zip -Zip "ABCDE"
} catch {
    $msg = $_.Exception.Message
}
if (-not ($msg.Contains("^\d{5}$") -or $msg.Contains("pattern") -or $msg.Contains("ABCDE"))) {
    Write-Host "FAIL: Error message should mention pattern and value, got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
