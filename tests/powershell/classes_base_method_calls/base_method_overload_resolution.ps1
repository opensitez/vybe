# vybe-test: powershell/classes_base_method_calls/base_method_overload_resolution
class BaseOverloader {
    [string]Show([int]$i) { return "Int:$i" }
    [string]Show([string]$s) { return "Str:$s" }
}
class SubOverloader : BaseOverloader {
    [string]ShowBoth() {
        $s1 = ([BaseOverloader]$this).Show(42)
        $s2 = ([BaseOverloader]$this).Show("test")
        return "$s1|$s2"
    }
}
$so = [SubOverloader]::new()
if ($so.ShowBoth() -ne "Int:42|Str:test") {
    Write-Host "FAIL: Base method overload resolution failed"
    exit 1
}
Write-Host "PASS"
exit 0
