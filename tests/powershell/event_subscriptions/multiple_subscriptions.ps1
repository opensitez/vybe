# vybe-test: powershell/event_subscriptions/multiple_subscriptions
Register-EngineEvent -SourceIdentifier MultiSub -Action { $Global.X += 1 }
Register-EngineEvent -SourceIdentifier MultiSub -Action { $Global.X += 2 }
New-Event -SourceIdentifier MultiSub
if ($Global.X -ne 3) {
    Write-Host "FAIL: expected 3 from two handlers"
    exit 1
}
Unregister-Event -SourceIdentifier MultiSub -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
