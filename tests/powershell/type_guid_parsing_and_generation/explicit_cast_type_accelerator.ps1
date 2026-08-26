# vybe-test: powershell/type_guid_parsing_and_generation/explicit_cast_type_accelerator
$str = "12345678-1234-1234-1234-123456789abc"
$g = [guid]$str
if ($g.GetType().Name -ne "Guid" -or $g.ToString() -ne $str) {
    Write-Host "FAIL: cast to [guid] failed"
    exit 1
}
Write-Host "PASS"
exit 0
