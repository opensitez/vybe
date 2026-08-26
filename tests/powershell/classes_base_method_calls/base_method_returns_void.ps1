# vybe-test: powershell/classes_base_method_calls/base_method_returns_void
class BaseLogger {
    [string]$LogData = ""
    [void]AppendLog([string]$s) { $this.LogData += "$s;" }
}
class ExtendedLogger : BaseLogger {
    [void]LogInfo([string]$info) {
        ([BaseLogger]$this).AppendLog("INFO:$info")
    }
}
$el = [ExtendedLogger]::new()
$el.LogInfo("Started")
$el.LogInfo("Done")
if ($el.LogData -ne "INFO:Started;INFO:Done;") {
    Write-Host "FAIL: Base method returning void failed, got '$($el.LogData)'"
    exit 1
}
Write-Host "PASS"
exit 0
