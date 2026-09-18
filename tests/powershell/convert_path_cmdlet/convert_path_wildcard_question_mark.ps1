# vybe-test: powershell/convert_path_cmdlet/convert_path_wildcard_question_mark
# Convert-Path resolves single-character wildcard '?' against matching paths
$res = Convert-Path "test?"

if ($null -eq $res) {
    Write-Host "FAIL: Convert-Path with '?' returned `$null"
    exit 1
}

if ($res -notmatch "tests$") {
    Write-Host "FAIL: expected 'test?' to resolve to 'tests', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
