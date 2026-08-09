# vybe-test: powershell/null_coalescing_assignment/null_assignment_hashtable_var
$map = $null
$map ??= @{ Status = "Init" }
if ($map.Status -ne "Init") {
    Write-Host "FAIL: hashtable variable ??= expected Status=Init"
    exit 1
}
Write-Host "PASS"
exit 0
