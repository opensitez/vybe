# vybe-test: powershell/classes_polymorphism_and_casting/polymorphic_event_notification_pattern
class Observer {
    [string]$LastMsg = ""
    [void]OnNotify([string]$msg) { $this.LastMsg = $msg }
}
class CustomObserver : Observer {
    [void]OnNotify([string]$msg) { $this.LastMsg = "CUSTOM:$msg" }
}
[Observer]$obs = [CustomObserver]::new()
$obs.OnNotify("event1")
if ($obs.LastMsg -ne "CUSTOM:event1") {
    Write-Host "FAIL: Polymorphic notification failed"
    exit 1
}
Write-Host "PASS"
exit 0
