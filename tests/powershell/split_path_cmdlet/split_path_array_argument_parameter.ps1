# vybe-test: powershell/split_path_cmdlet/split_path_array_argument_parameter
# Split-Path accepts an array of paths via the -Path parameter
$results = Split-Path -Path @("/system/core.dll", "/system/kernel.dll") -Leaf

if ($results.Count -ne 2) {
    Write-Host "FAIL: expected 2 results, got $($results.Count)"
    exit 1
}

if ($results[0] -ne "core.dll" -or $results[1] -ne "kernel.dll") {
    Write-Host "FAIL: array leaves mismatch, got: @($($results -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
