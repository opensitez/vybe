# vybe-test: powershell/classes_this_chaining_and_scope/this_accessed_inside_scriptblock_inside_method
class Aggregator {
    [int]$Multiplier = 3
    [int[]]ProcessArray([int[]]$arr) {
        $mult = $this.Multiplier
        return @($arr | ForEach-Object { $_ * $mult })
    }
}
$ag = [Aggregator]::new()
$res = $ag.ProcessArray(@(1, 2, 3))
if ($res[0] -ne 3 -or $res[1] -ne 6 -or $res[2] -ne 9) {
    Write-Host "FAIL: `$this in method scriptblock failed"
    exit 1
}
Write-Host "PASS"
exit 0
