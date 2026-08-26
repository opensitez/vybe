# vybe-test: powershell/classes_base_method_calls/call_base_method_via_cast
class BaseGreeter {
    [string]Greet() { return "Hello from Base" }
}
class SubGreeter : BaseGreeter {
    [string]Greet() {
        $baseMsg = ([BaseGreeter]$this).Greet()
        return "$baseMsg and Sub"
    }
}
$sg = [SubGreeter]::new()
$msg = $sg.Greet()
if ($msg -ne "Hello from Base and Sub") {
    Write-Host "FAIL: Base method call via cast failed, got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
