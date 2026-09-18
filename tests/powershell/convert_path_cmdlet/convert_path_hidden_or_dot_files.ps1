# vybe-test: powershell/convert_path_cmdlet/convert_path_hidden_or_dot_files
# Convert-Path correctly resolves hidden dot-directories like .git
$res = Convert-Path ".git"

if ($null -eq $res) {
    Write-Host "FAIL: Convert-Path returned `$null for .git"
    exit 1
}

if ($res -notmatch "\.git$") {
    Write-Host "FAIL: expected path to end with '.git', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
