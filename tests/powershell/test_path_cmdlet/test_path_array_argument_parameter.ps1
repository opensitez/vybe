# vybe-test: powershell/test_path_cmdlet/test_path_array_argument_parameter
# Passing an array of paths to -Path returns an array of booleans
$results = Test-Path -Path @("tests", "Cargo.toml")

if ($results.Count -ne 2) {
    Write-Host "FAIL: expected 2 results, got $($results.Count)"
    exit 1
}

if ($results[0] -ne $true -or $results[1] -ne $true) {
    Write-Host "FAIL: expected @(`$true, `$true), got: @($($results -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
