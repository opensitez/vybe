# vybe-test: powershell/classes/class_method_overload
class Formatter {
    [string]Format([int]$n)         { return "int:$n" }
    [string]Format([double]$d)      { return "double:$d" }
    [string]Format([string]$s)      { return "string:$s" }
}
$f = [Formatter]::new()
if ($f.Format(42)     -ne "int:42")      { Write-Host "FAIL: int overload";    exit 1 }
if ($f.Format("hi")   -ne "string:hi")   { Write-Host "FAIL: string overload"; exit 1 }
Write-Host "PASS"
exit 0
