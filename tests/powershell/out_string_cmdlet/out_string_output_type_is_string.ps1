# vybe-test: powershell/out_string_cmdlet/out_string_output_type_is_string
# Out-String strictly emits System.String instances
$output = (1..10) | Out-String

if (-not ($output -is [string])) {
    Write-Host "FAIL: expected [string], got: $($output.GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
