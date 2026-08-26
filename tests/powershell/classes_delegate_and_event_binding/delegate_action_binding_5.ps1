# vybe-test: powershell/classes_delegate_and_event_binding/delegate_action_binding_5
$action = [System.Action[int]]{ param($n) $global:actionResult = $n * 3 }
$action.Invoke(5)
if ($global:actionResult -ne (5 * 3)) { Write-Host "FAIL: Delegate action binding failed"; exit 1 }
Write-Host "PASS"; exit 0
