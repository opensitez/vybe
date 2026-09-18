# vybe-test: powershell/select_string_cmdlet/select_string_ignorecase_property
# MatchInfo exposes .IgnoreCase reflecting whether -CaseSensitive was specified
$matchDefault = "sample" | Select-String -Pattern "sample"
$matchStrict  = "sample" | Select-String -Pattern "sample" -CaseSensitive

if ($matchDefault.IgnoreCase -ne $true) {
    Write-Host "FAIL: default .IgnoreCase expected `$true, got: $($matchDefault.IgnoreCase)"
    exit 1
}

if ($matchStrict.IgnoreCase -ne $false) {
    Write-Host "FAIL: with -CaseSensitive .IgnoreCase expected `$false, got: $($matchStrict.IgnoreCase)"
    exit 1
}

Write-Host "PASS"
exit 0
