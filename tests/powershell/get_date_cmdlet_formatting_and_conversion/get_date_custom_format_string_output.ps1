# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_custom_format_string_output
# Get-Date -Format evaluates standard and custom .NET DateTime format strings into string output
$date = [DateTime]::Parse("2026-05-10 14:30:45")
$formatted = Get-Date -Date $date -Format "yyyy_MM_dd"

if (-not ($formatted -is [string])) {
    Write-Host "FAIL: expected string return type with -Format, got: $($formatted.GetType().FullName)"
    exit 1
}

if ($formatted -ne "2026_05_10") {
    Write-Host "FAIL: formatted string mismatch, expected '2026_05_10', got: '$formatted'"
    exit 1
}

Write-Host "PASS"
exit 0
