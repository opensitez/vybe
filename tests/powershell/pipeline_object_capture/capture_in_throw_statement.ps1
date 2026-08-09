# vybe-test: powershell/pipeline_object_capture/capture_in_throw_statement
try {
    throw (1..3 | ForEach-Object { "ERR_$_" }) -join ":"
} catch {
    if ($_.Exception.Message -ne "ERR_1:ERR_2:ERR_3") {
        Write-Host "FAIL: throw statement with pipeline join expected ERR_1:ERR_2:ERR_3, got '$($_.Exception.Message)'"
        exit 1
    }
}
Write-Host "PASS"
exit 0
