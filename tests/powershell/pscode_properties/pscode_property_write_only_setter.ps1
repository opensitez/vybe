# vybe-test: powershell/pscode_properties/pscode_property_write_only_setter
class LogSinkHelper {
    static [string]$LastMsg = ""
    static [void] SetMsg([object]$t, [string]$msg) { [LogSinkHelper]::$LastMsg = $msg }
}
$s = [LogSinkHelper].GetMethod("SetMsg")
$cp = [System.Management.Automation.PSCodeProperty]::new("Sink", $null, $s)
if ($cp.IsGettable -or -not $cp.IsSettable) {
    Write-Host "FAIL: write-only PSCodeProperty expected IsGettable=false, IsSettable=true"
    exit 1
}
Write-Host "PASS"
exit 0
