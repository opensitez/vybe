# vybe-test: powershell/type_accelerators/type_accelerator_guid
$str = "d3b07384-d113-40a6-a1e4-1050974b88fe"
$g = [guid]$str
if ($g.ToString() -ne $str) {
    Write-Host "FAIL: Guid string expected $str, got $($g.ToString())"
    exit 1
}
Write-Host "PASS"
exit 0
