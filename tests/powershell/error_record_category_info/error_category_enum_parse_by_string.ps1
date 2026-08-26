# vybe-test: powershell/error_record_category_info/error_category_enum_parse_by_string
$cat = [System.Enum]::Parse([System.Management.Automation.ErrorCategory], "AuthenticationError")
if ($cat -ne [System.Management.Automation.ErrorCategory]::AuthenticationError) {
    Write-Host "FAIL: ErrorCategory parse by string failed"
    exit 1
}
Write-Host "PASS"
exit 0
