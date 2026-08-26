# vybe-test: powershell/exceptions_custom_net_exception_classes/custom_exception_with_enum_severity
enum SeverityLevel { Low; Medium; High; Critical }
class SeverityException : System.Exception {
    [SeverityLevel]$Severity
    SeverityException([SeverityLevel]$lvl, [string]$msg) : base($msg) {
        $this.Severity = $lvl
    }
}
$se = [SeverityException]::new([SeverityLevel]::Critical, "Database down")
if ($se.Severity -ne [SeverityLevel]::Critical) {
    Write-Host "FAIL: Custom exception with enum severity failed"
    exit 1
}
Write-Host "PASS"
exit 0
