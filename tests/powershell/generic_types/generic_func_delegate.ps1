# vybe-test: powershell/generic_types/generic_func_delegate
$func = [Func[int, int, int]]{ param($a, $b) $a + $b }
$val = $func.Invoke(20, 22)
if ($val -ne 42) {
    Write-Host "FAIL: Func[int,int,int] expected 42, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
