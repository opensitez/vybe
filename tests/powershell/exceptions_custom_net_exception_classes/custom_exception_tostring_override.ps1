# vybe-test: powershell/exceptions_custom_net_exception_classes/custom_exception_tostring_override
class PrettyException : System.Exception {
    [int]$ErrCode
    PrettyException([int]$code, [string]$m) : base($m) { $this.ErrCode = $code }
    [string]ToString() {
        return "ERROR-$($this.ErrCode): $($this.Message)"
    }
}
$pe = [PrettyException]::new(500, "Server Crash")
if ($pe.ToString() -ne "ERROR-500: Server Crash") {
    Write-Host "FAIL: Custom exception ToString override failed, got '$($pe.ToString())'"
    exit 1
}
Write-Host "PASS"
exit 0
