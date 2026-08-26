# vybe-test: powershell/exceptions_throw_expression_vs_statement/throw_error_record_instance_preserves_error_record
$ex = [System.InvalidOperationException]::new("Custom")
$errRec = [System.Management.Automation.ErrorRecord]::new($ex, "CustomId", [System.Management.Automation.ErrorCategory]::InvalidOperation, $null)
$caught = $null
try {
    throw $errRec
} catch {
    $caught = $_
}
if ($caught.FullyQualifiedErrorId -ne "CustomId") {
    Write-Host "FAIL: Throw ErrorRecord preservation failed, got '$($caught.FullyQualifiedErrorId)'"
    exit 1
}
Write-Host "PASS"
exit 0
