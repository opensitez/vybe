# vybe-test: powershell/error_record_category_info/write_error_with_category_parameter
function Test-CategoryEmit {
    [CmdletBinding()]
    param()
    Write-Error -Message "Access rejected" -Category PermissionDenied
}
$errRecord = $null
try {
    Test-CategoryEmit -ErrorAction Stop
} catch {
    $errRecord = $_
}
if ($errRecord.CategoryInfo.Category -ne [System.Management.Automation.ErrorCategory]::PermissionDenied) {
    Write-Host "FAIL: Write-Error -Category check failed, got $($errRecord.CategoryInfo.Category)"
    exit 1
}
Write-Host "PASS"
exit 0
