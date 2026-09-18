# vybe-test: powershell/out_string_cmdlet/out_string_error_record_formatting
# Out-String converts trapped ErrorRecord instances into diagnostic error strings
$errRecord = $null
try {
    1 / 0
} catch {
    $errRecord = $_
}

if ($null -eq $errRecord) {
    Write-Host "FAIL: exception was not trapped"
    exit 1
}

$output = $errRecord | Out-String

if ($output -notmatch "Attempted to divide by zero" -and $output -notmatch "DivideByZero") {
    Write-Host "FAIL: expected divide by zero diagnostic message in Out-String output, got: '$output'"
    exit 1
}

Write-Host "PASS"
exit 0
