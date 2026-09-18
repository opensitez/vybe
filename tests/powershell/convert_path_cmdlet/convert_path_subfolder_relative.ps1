# vybe-test: powershell/convert_path_cmdlet/convert_path_subfolder_relative
# Convert-Path on a relative subfolder path returns the full absolute path
$res = Convert-Path "tests"

$expectedUnix = "$($PWD.Path)/tests"
$expectedWin  = "$($PWD.Path)\tests"

if ($res -ne $expectedUnix -and $res -ne $expectedWin) {
    Write-Host "FAIL: relative subfolder conversion failed, expected '$expectedUnix', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
