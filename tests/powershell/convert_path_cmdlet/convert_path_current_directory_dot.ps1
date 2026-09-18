# vybe-test: powershell/convert_path_cmdlet/convert_path_current_directory_dot
# Convert-Path on current directory '.' returns the absolute filesystem path
$res = Convert-Path "."

if ($res -ne $PWD.Path) {
    Write-Host "FAIL: expected current directory '$($PWD.Path)', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
