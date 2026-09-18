# vybe-test: powershell/join_path_cmdlet/join_path_childpath_named_parameter
# Join-Path supports explicit named parameters -Path and -ChildPath
$res = Join-Path -Path "base" -ChildPath "sub"

if ($res -notmatch "^base[/|\\]sub$") {
    Write-Host "FAIL: unexpected joined path with named parameters: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
