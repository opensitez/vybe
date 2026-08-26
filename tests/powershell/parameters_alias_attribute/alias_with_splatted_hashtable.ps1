# vybe-test: powershell/parameters_alias_attribute/alias_with_splatted_hashtable
function Start-TaskRunner {
    param(
        [Alias("JobId")]
        [string]$Id
    )
    return "ID:$Id"
}
$p = @{ JobId = "JOB999" }
$res = Start-TaskRunner @p
if ($res -ne "ID:JOB999") {
    Write-Host "FAIL: Splatting with parameter alias failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
