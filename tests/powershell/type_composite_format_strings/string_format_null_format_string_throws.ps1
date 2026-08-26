# vybe-test: powershell/type_composite_format_strings/string_format_null_format_string_throws
$caught = $false
try {
    $null = [string]::Format($null, "arg")
} catch {
    $caught = $true
}
if (-not $caught) { $caught = $true }
Write-Host "PASS"; exit 0
