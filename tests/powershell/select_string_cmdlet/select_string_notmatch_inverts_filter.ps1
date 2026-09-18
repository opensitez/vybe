# vybe-test: powershell/select_string_cmdlet/select_string_notmatch_inverts_filter
# -NotMatch filters out matching lines and retains only lines that do not match the pattern
$stream = @("keep_alpha", "drop_beta", "keep_gamma")
$filtered = @($stream | Select-String -Pattern "drop" -NotMatch)

if ($filtered.Count -ne 2) {
    Write-Host "FAIL: expected 2 non-matching lines, got $($filtered.Count)"
    exit 1
}

$retained = @($filtered | ForEach-Object { $_.Line })
if ($retained[0] -ne "keep_alpha" -or $retained[1] -ne "keep_gamma") {
    Write-Host "FAIL: unexpected retained lines: @($($retained -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
