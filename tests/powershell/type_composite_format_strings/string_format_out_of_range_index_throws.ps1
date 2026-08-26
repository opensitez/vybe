# vybe-test: powershell/type_composite_format_strings/string_format_out_of_range_index_throws
$caught = $false
try {
    $x = [string]::Format("{0} and {1}", "One")
} catch [System.FormatException] {
    $caught = $true
}
if (-not $caught) { Write-Host "FAIL: FormatException expected on missing index"; exit 1 }
Write-Host "PASS"; exit 0
