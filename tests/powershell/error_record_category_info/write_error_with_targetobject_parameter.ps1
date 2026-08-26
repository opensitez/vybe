# vybe-test: powershell/error_record_category_info/write_error_with_targetobject_parameter
function Test-TargetEmit {
    [CmdletBinding()]
    param()
    Write-Error -Message "Missing item" -TargetObject "/etc/config.json" -Category ObjectNotFound
}
$errRecord = $null
try {
    Test-TargetEmit -ErrorAction Stop
} catch {
    $errRecord = $_
}
if ($errRecord.TargetObject -ne "/etc/config.json") {
    Write-Host "FAIL: Write-Error -TargetObject check failed"
    exit 1
}
Write-Host "PASS"
exit 0
