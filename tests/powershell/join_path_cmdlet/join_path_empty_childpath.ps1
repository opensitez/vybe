# vybe-test: powershell/join_path_cmdlet/join_path_empty_childpath
# Joining an empty ChildPath string returns the base path without crashing
$res = Join-Path -Path "base" -ChildPath ""

if ($res -ne "base" -and $res -ne "base/" -and $res -ne "base\") {
    Write-Host "FAIL: unexpected result for empty ChildPath: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
