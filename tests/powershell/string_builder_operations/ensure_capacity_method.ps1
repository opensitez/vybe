# vybe-test: powershell/string_builder_operations/ensure_capacity_method
$sb = [System.Text.StringBuilder]::new()
$cap = $sb.EnsureCapacity(128)
if ($cap -lt 128) {
    Write-Host "FAIL: EnsureCapacity failed, got $cap"
    exit 1
}
Write-Host "PASS"
exit 0
