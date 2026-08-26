# vybe-test: powershell/error_record_category_info/categoryinfo_tostring_format
$ex = [System.Exception]::new("Err")
$err = [System.Management.Automation.ErrorRecord]::new(
    $ex,
    "ErrId",
    [System.Management.Automation.ErrorCategory]::SyntaxError,
    "code.ps1"
)
$str = $err.CategoryInfo.ToString()
if (-not ($str.Contains("SyntaxError") -and $str.Contains("code.ps1"))) {
    Write-Host "FAIL: CategoryInfo ToString format failed, got '$str'"
    exit 1
}
Write-Host "PASS"
exit 0
