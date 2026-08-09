# vybe-test: powershell/type_converters/type_converter_guid_string
$gStr = "00000000-0000-0000-0000-000000000000"
$g = [guid]$gStr
if ($g -ne [guid]::Empty) {
    Write-Host "FAIL: string to [guid] empty conversion expected empty guid"
    exit 1
}
Write-Host "PASS"
exit 0
