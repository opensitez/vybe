# vybe-test: powershell/event_subscriptions/subscription_message_data
Register-EngineEvent -SourceIdentifier DataSub -Action { param($e) $Global.Data = $e.MessageData }
New-Event -SourceIdentifier DataSub -MessageData 42
if ($Global.Data -ne 42) {
    Write-Host "FAIL: expected message data 42"
    exit 1
}
Unregister-Event -SourceIdentifier DataSub -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
