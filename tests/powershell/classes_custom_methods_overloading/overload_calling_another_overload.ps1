# vybe-test: powershell/classes_custom_methods_overloading/overload_calling_another_overload
class LoggerDelegator {
    [string]Log([string]$msg) {
        return $this.Log($msg, "INFO")
    }
    [string]Log([string]$msg, [string]$level) {
        return "[$level] $msg"
    }
}
$ld = [LoggerDelegator]::new()
$res = $ld.Log("Server started")
if ($res -ne "[INFO] Server started") {
    Write-Host "FAIL: Overload delegating to another overload failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
