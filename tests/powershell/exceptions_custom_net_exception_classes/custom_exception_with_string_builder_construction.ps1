# vybe-test: powershell/exceptions_custom_net_exception_classes/custom_exception_with_string_builder_construction
class FormattedReportException : System.Exception {
    FormattedReportException([string[]]$reasons) : base(
        ([System.String]::Join("; ", $reasons))
    ) {}
}
$fre = [FormattedReportException]::new(@("R1", "R2", "R3"))
if ($fre.Message -ne "R1; R2; R3") {
    Write-Host "FAIL: Custom exception message formatting failed, got '$($fre.Message)'"
    exit 1
}
Write-Host "PASS"
exit 0
