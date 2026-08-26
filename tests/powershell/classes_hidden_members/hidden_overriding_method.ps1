# vybe-test: powershell/classes_hidden_members/hidden_overriding_method
class BaseWorker {
    [string]Work() { return "base" }
}
class FastWorker : BaseWorker {
    hidden [string]Work() { return "fast" }
    [string]Run() { return $this.Work() }
}
$fw = [FastWorker]::new()
if ($fw.Run() -ne "fast") {
    Write-Host "FAIL: Hidden overriding method failed"
    exit 1
}
Write-Host "PASS"
exit 0
