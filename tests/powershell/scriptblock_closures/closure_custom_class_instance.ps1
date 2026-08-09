# vybe-test: powershell/scriptblock_closures/closure_custom_class_instance
class ConfigObj {
    [string]$Mode = "Fast"
}
$cfg = [ConfigObj]::new()
$sb = { $cfg.Mode }.GetClosure()
$res = &$sb
if ($res -ne "Fast") {
    Write-Host "FAIL: custom class instance capture in closure expected 'Fast', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
