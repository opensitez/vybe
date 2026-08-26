# vybe-test: powershell/error_record_exception_chaining/exception_chain_iteration_loop
$e1 = [System.Exception]::new("1")
$e2 = [System.Exception]::new("2", $e1)
$e3 = [System.Exception]::new("3", $e2)
$collected = [System.Collections.Generic.List[string]]::new()
$cur = $e3
while ($cur -ne $null) {
    $collected.Add($cur.Message)
    $cur = $cur.InnerException
}
if ($collected.Count -ne 3 -or $collected[0] -ne "3" -or $collected[2] -ne "1") {
    Write-Host "FAIL: Exception chain iteration loop failed"
    exit 1
}
Write-Host "PASS"
exit 0
