# vybe-test: powershell/convert_path_cmdlet/convert_path_array_argument_parameter
# Passing an array to the -Path parameter converts all specified paths into a string array
$results = Convert-Path -Path @(".", "tests")

if ($results.Count -ne 2) {
    Write-Host "FAIL: expected 2 results, got $($results.Count)"
    exit 1
}

if ($results[0] -ne $PWD.Path -or ($results[1] -notmatch "tests$")) {
    Write-Host "FAIL: array argument conversion mismatch: @($($results -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
