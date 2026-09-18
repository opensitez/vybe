# vybe-test: powershell/convert_path_cmdlet/convert_path_trailing_slash_preservation
# In PowerShell Convert-Path, supplying a trailing slash on a directory path preserves the trailing separator
$res = Convert-Path "tests/"
$baseExpected = Convert-Path "tests"

if ($res -ne "$baseExpected/" -and $res -ne "$baseExpected\") {
    Write-Host "FAIL: trailing slash was not preserved in converted path, expected '$baseExpected/', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
