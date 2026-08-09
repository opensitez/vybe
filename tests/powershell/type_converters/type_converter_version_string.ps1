# vybe-test: powershell/type_converters/type_converter_version_string
$ver = [version]"5.1.0.0"
if ($ver.Major -ne 5 -or $ver.Minor -ne 1) {
    Write-Host "FAIL: string to [version] conversion expected 5.1"
    exit 1
}
Write-Host "PASS"
exit 0
