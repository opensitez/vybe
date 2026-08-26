# vybe-test: powershell/error_record_category_info/error_record_in_try_catch_under_dollar_under
$caughtCat = $null
try {
    [int]::Parse("not-a-number")
} catch {
    $caughtCat = $_.CategoryInfo.Category
}
if ($caughtCat -ne [System.Management.Automation.ErrorCategory]::InvalidArgument -and $caughtCat -ne [System.Management.Automation.ErrorCategory]::NotSpecified) {
    if ($caughtCat -eq $null) {
        Write-Host "FAIL: `$_ in catch block missing CategoryInfo"
        exit 1
    }
}
Write-Host "PASS"
exit 0
