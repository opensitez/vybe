# vybe-test: powershell/classes_constructor_overloading/constructor_null_argument_dispatch
class NullDispatchTarget {
    [string]$Type
    NullDispatchTarget([string]$s) { $this.Type = "String" }
    NullDispatchTarget([hashtable]$h) { $this.Type = "Hashtable" }
}
$t1 = [NullDispatchTarget]::new("hello")
$t2 = [NullDispatchTarget]::new(@{ a = 1 })
if ($t1.Type -ne "String" -or $t2.Type -ne "Hashtable") {
    Write-Host "FAIL: Constructor typed dispatch failed"
    exit 1
}
Write-Host "PASS"
exit 0
