# vybe-test: powershell/exceptions_finally_guarantees_with_return/finally_in_class_method_with_return
class MethodFinallyTarget {
    [bool]$Cleaned = $false
    [string]Run() {
        try {
            return "SUCCESS"
        } finally {
            $this.Cleaned = $true
        }
    }
}
$mft = [MethodFinallyTarget]::new()
$ret = $mft.Run()
if ($ret -ne "SUCCESS" -or -not $mft.Cleaned) {
    Write-Host "FAIL: Finally in class method failed"
    exit 1
}
Write-Host "PASS"
exit 0
