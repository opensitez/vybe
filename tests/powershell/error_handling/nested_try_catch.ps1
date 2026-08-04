# vybe-test: powershell/error_handling/nested_try_catch
$log = @()
try {
    try {
        throw "inner error"
    } catch {
        $log += "inner caught"
        throw "rethrown"
    }
} catch {
    $log += "outer caught: $($_.Exception.Message)"
}
if ($log[0] -ne "inner caught")               { Write-Host "FAIL: inner"; exit 1 }
if ($log[1] -ne "outer caught: rethrown")     { Write-Host "FAIL: outer '$($log[1])'"; exit 1 }
Write-Host "PASS"
exit 0
