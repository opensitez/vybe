# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_pipeline_datetime_input_formatting
# Streaming DateTime objects through the pipeline allows formatting via Get-Date parameters
$date = [DateTime]::Parse("2026-05-10 14:30:45")
$formatted = $date | Get-Date -Format "dd/MM/yyyy"

if ($formatted -ne "10/05/2026") {
    Write-Host "FAIL: pipeline formatting mismatch, expected '10/05/2026', got: '$formatted'"
    exit 1
}

Write-Host "PASS"
exit 0
