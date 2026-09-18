# vybe-test: powershell/test_path_cmdlet/test_path_isvalid_parameter_syntax_check
# -IsValid checks whether the path syntax is well-formed, regardless of whether the path exists
$res = Test-Path "subfolder_that_does_not_exist/file_xyz.txt" -IsValid

if ($res -ne $true) {
    Write-Host "FAIL: -IsValid returned `$false for syntactically valid path"
    exit 1
}

Write-Host "PASS"
exit 0
