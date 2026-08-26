# vybe-test: powershell/error_record_category_info/write_error_recommended_action
function Test-ActionEmit {
    [CmdletBinding()]
    param()
    Write-Error -Message "Timeout" -RecommendedAction "Check network connection"
}
$err = $null
try {
    Test-ActionEmit -ErrorAction Stop
} catch {
    $err = $_
}
if ($err.ErrorDetails.RecommendedAction -ne "Check network connection") {
    Write-Host "FAIL: RecommendedAction check failed"
    exit 1
}
Write-Host "PASS"
exit 0
