# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_target_object_preserved
# ErrorRecord.TargetObject and CategoryInfo.TargetName preserve the target object supplied in constructor
function EmitTargetObjectCheck {
    [CmdletBinding()]
    param()
    process {
        $target = "redis.internal:6379"
        $ex = [System.TimeoutException]::new("Redis ping timed out")
        $err = [System.Management.Automation.ErrorRecord]::new(
            $ex,
            "RedisTimeoutId",
            [System.Management.Automation.ErrorCategory]::OperationTimeout,
            $target
        )
        $PSCmdlet.WriteError($err)
    }
}

$captured = @(EmitTargetObjectCheck 2>&1)
$errRecord = $captured[0]

if ($errRecord.TargetObject -ne "redis.internal:6379") {
    Write-Host "FAIL: TargetObject mismatch: '$($errRecord.TargetObject)'"
    exit 1
}

if ($errRecord.CategoryInfo.TargetName -ne "redis.internal:6379") {
    Write-Host "FAIL: CategoryInfo.TargetName mismatch: '$($errRecord.CategoryInfo.TargetName)'"
    exit 1
}

Write-Host "PASS"
exit 0
