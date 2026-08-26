# vybe-test: powershell/error_record_category_info/error_record_fullyqualifiederrorid_property
$ex = [System.Exception]::new("Err")
$err = [System.Management.Automation.ErrorRecord]::new($ex, "CustomId", [System.Management.Automation.ErrorCategory]::NotSpecified, $null)
if ($err.FullyQualifiedErrorId -ne "CustomId") {
    Write-Host "FAIL: FullyQualifiedErrorId check failed, got '$($err.FullyQualifiedErrorId)'"
    exit 1
}
Write-Host "PASS"
exit 0
