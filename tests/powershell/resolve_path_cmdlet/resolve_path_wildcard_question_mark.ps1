# vybe-test: powershell/resolve_path_cmdlet/resolve_path_wildcard_question_mark
# Resolve-Path interprets '?' as matching exactly one character
$info = Resolve-Path "test?"

if ($null -eq $info) {
    Write-Host "FAIL: Resolve-Path with '?' returned `$null"
    exit 1
}

if ($info.Path -notmatch "tests$") {
    Write-Host "FAIL: expected 'test?' to resolve to 'tests', got: '$($info.Path)'"
    exit 1
}

Write-Host "PASS"
exit 0
