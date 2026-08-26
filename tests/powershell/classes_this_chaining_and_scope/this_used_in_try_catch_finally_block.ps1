# vybe-test: powershell/classes_this_chaining_and_scope/this_used_in_try_catch_finally_block
class SafeOp {
    [bool]$Cleaned = $false
    [string]Execute() {
        try {
            return "SUCCESS"
        } finally {
            $this.Cleaned = $true
        }
    }
}
$so = [SafeOp]::new()
$res = $so.Execute()
if ($res -ne "SUCCESS" -or -not $so.Cleaned) {
    Write-Host "FAIL: `$this in finally block failed"
    exit 1
}
Write-Host "PASS"
exit 0
