# vybe-test: powershell/generic_types/generic_action_delegate
$res = 0
$act = [Action[int]]{ param($x) $script:res = $x * 10 }
$act.Invoke(5)
if ($res -ne 50) {
    Write-Host "FAIL: Action[int] delegate invocation expected 50, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
