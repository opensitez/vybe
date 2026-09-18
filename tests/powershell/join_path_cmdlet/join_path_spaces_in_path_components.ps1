# vybe-test: powershell/join_path_cmdlet/join_path_spaces_in_path_components
# Spaces within directory or file components must be preserved across joins
$res = Join-Path "Program Files" "My App" "config.ini"

if ($res -notmatch "Program Files[/|\\]My App[/|\\]config\.ini") {
    Write-Host "FAIL: path with spaces failed to join properly: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
