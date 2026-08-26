# vybe-test: powershell/classes_constructor_overloading/constructor_differing_by_type
class ValueHolder {
    [string]$Type
    [string]$Value
    ValueHolder([int]$i) {
        $this.Type = "Int"
        $this.Value = "$i"
    }
    ValueHolder([string]$s) {
        $this.Type = "String"
        $this.Value = $s
    }
}
$h1 = [ValueHolder]::new(42)
$h2 = [ValueHolder]::new("hello")
if ($h1.Type -ne "Int" -or $h2.Type -ne "String") {
    Write-Host "FAIL: Constructor type dispatch failed"
    exit 1
}
Write-Host "PASS"
exit 0
