# vybe-test: powershell/whitespace_and_line_rules/carriage_return_in_string_literals
$payload = "line1`rline2"
if (-not $payload.Contains("`r")) {
    Write-Host 'FAIL: carriage return missing from quoted string'
    exit 1
}

if ($payload -ne "line1`rline2") {
    Write-Host 'FAIL: carriage return content altered'
    exit 1
}

Write-Host 'PASS'
exit 0
