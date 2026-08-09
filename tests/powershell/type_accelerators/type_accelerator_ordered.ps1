# vybe-test: powershell/type_accelerators/type_accelerator_ordered
$ord = [ordered]@{ First = 1; Second = 2; Third = 3 }
$keys = @($ord.Keys)
if ($keys[0] -ne "First" -or $keys[1] -ne "Second" -or $keys[2] -ne "Third") {
    Write-Host "FAIL: ordered hashtable key order not preserved, got $($keys -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
