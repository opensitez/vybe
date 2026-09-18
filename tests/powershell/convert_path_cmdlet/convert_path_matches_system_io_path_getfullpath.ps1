# vybe-test: powershell/convert_path_cmdlet/convert_path_matches_system_io_path_getfullpath
# Convert-Path on filesystem items produces canonical paths matching .NET System.IO.Path.GetFullPath
$cmdletResult = Convert-Path "tests"
$dotnetResult = [System.IO.Path]::GetFullPath("tests")

if ($cmdletResult -ne $dotnetResult) {
    Write-Host "FAIL: Convert-Path ('$cmdletResult') did not match [System.IO.Path]::GetFullPath ('$dotnetResult')"
    exit 1
}

Write-Host "PASS"
exit 0
