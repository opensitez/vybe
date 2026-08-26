# vybe-test: powershell/classes_delegate_and_event_binding/delegate_action_binding_14
$action = [System.Action[int]]{ param($n) $global:actionResult = $n * 3 }
$action.Invoke(14)
if ($global:actionResult -ne (14 * 3)) { Write-Host "FAIL: Delegate action binding failed"; exit 1 }
Write-Host "PASS"; exit 0
