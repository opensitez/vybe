# vybe-test: powershell/exceptions_try_catch_multiple_typed_blocks/catch_specific_guid_parse_format_exception
$caught = $false
try {
    $g = [guid]::Parse("not-a-guid")
} catch [System.FormatException] {
    $caught = $true
} catch {
    $caught = $false
}
if (-not $caught) {
    Write-Host "FAIL: Catching GUID FormatException failed"
    exit 1
}
Write-Host "PASS"
exit 0
