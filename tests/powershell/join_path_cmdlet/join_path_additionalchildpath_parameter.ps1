# vybe-test: powershell/join_path_cmdlet/join_path_additionalchildpath_parameter
# -AdditionalChildPath accepts an array of child segments to append
$res = Join-Path -Path "base" -ChildPath "sub" -AdditionalChildPath @("sub2", "file.txt")

if ($res -notmatch "^base[/|\\]sub[/|\\]sub2[/|\\]file\.txt$") {
    Write-Host "FAIL: -AdditionalChildPath joining failed: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
