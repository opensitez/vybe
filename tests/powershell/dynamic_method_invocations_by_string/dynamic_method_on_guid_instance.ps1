# vybe-test: powershell/dynamic_method_invocations_by_string/dynamic_method_on_guid_instance
$g = [guid]::NewGuid()
$m = "ToByteArray"
$bytes = $g.$m()
if ($bytes.Length -ne 16) {
    Write-Host "FAIL: Dynamic method on GUID failed"
    exit 1
}
Write-Host "PASS"
exit 0
