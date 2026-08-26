# vybe-test: powershell/classes_polymorphism_and_casting/downcast_inside_function
class BaseJob { [string]$Title; BaseJob([string]$t) { $this.Title = $t } }
class UrgentJob : BaseJob { [int]$Priority = 1; UrgentJob([string]$t) : base($t) {} }
function Get-JobPriority([BaseJob]$j) {
    if ($j -is [UrgentJob]) {
        return ([UrgentJob]$j).Priority
    }
    return 0
}
$uj = [UrgentJob]::new("FixBug")
$bj = [BaseJob]::new("Normal")
if ((Get-JobPriority $uj) -ne 1 -or (Get-JobPriority $bj) -ne 0) {
    Write-Host "FAIL: Downcast inside function failed"
    exit 1
}
Write-Host "PASS"
exit 0
