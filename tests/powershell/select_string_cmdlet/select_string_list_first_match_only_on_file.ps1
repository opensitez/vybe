# vybe-test: powershell/select_string_cmdlet/select_string_list_first_match_only_on_file
# -List stops after the first matching line in a file and returns exactly one MatchInfo
$matches = Select-String -Path "Cargo.toml" -Pattern "workspace" -List

if ($null -eq $matches) {
    Write-Host "FAIL: Select-String -List returned `$null"
    exit 1
}

if ($matches.Count -ne 1) {
    Write-Host "FAIL: expected exactly 1 match with -List on file, got: $($matches.Count)"
    exit 1
}

if ($matches.LineNumber -ne 1) {
    Write-Host "FAIL: expected first match on line 1, got line $($matches.LineNumber)"
    exit 1
}

Write-Host "PASS"
exit 0
