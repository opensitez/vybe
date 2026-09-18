# vybe-test: powershell/join_path_cmdlet/join_path_direct_named_parameters
# Join-Path explicitly binding -Path and -ChildPath
$res = Join-Path -Path "system" -ChildPath "drivers"

if ($res -notmatch "^system[/|\\]drivers$") {
    Write-Host "FAIL: direct named parameters failed: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
