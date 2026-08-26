# vybe-test: powershell/classes_constructor_overloading/constructor_with_timespan_parameter
class CacheEntry {
    [timespan]$Ttl
    CacheEntry([timespan]$t) { $this.Ttl = $t }
}
$ce = [CacheEntry]::new([timespan]::FromMinutes(5))
if ($ce.Ttl.TotalMinutes -ne 5.0) {
    Write-Host "FAIL: TimeSpan constructor failed"
    exit 1
}
Write-Host "PASS"
exit 0
