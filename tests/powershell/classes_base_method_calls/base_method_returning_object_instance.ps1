# vybe-test: powershell/classes_base_method_calls/base_method_returning_object_instance
class ResultWrapper {
    [string]$Status
    ResultWrapper([string]$s) { $this.Status = $s }
}
class BaseService {
    [ResultWrapper]Execute() { return [ResultWrapper]::new("BaseSuccess") }
}
class SubService : BaseService {
    [string]Run() {
        $res = ([BaseService]$this).Execute()
        return $res.Status
    }
}
$ss = [SubService]::new()
if ($ss.Run() -ne "BaseSuccess") {
    Write-Host "FAIL: Base method returning object instance failed"
    exit 1
}
Write-Host "PASS"
exit 0
