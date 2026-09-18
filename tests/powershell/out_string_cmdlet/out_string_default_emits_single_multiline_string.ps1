# vybe-test: powershell/out_string_cmdlet/out_string_default_emits_single_multiline_string
# By default, Out-String aggregates all input objects into a single multi-line string
$lines = @("first line", "second line")
$output = $lines | Out-String

if ($output.GetType().FullName -ne "System.String") {
    Write-Host "FAIL: output type is not [string], got: $($output.GetType().FullName)"
    exit 1
}

if ($output -notmatch "`n") {
    Write-Host "FAIL: expected newline characters in aggregated string, got: '$output'"
    exit 1
}

Write-Host "PASS"
exit 0
