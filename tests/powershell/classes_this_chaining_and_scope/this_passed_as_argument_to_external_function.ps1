# vybe-test: powershell/classes_this_chaining_and_scope/this_passed_as_argument_to_external_function
class TaskItem {
    [string]$Name
    TaskItem([string]$n) { $this.Name = $n }
    [string]GetFormatted() {
        return Format-Task $this
    }
}
function Format-Task([TaskItem]$t) {
    return "TASK:$($t.Name)"
}
$ti = [TaskItem]::new("Backup")
$res = $ti.GetFormatted()
if ($res -ne "TASK:Backup") {
    Write-Host "FAIL: Passing `$this to function failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
