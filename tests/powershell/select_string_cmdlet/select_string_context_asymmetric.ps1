# vybe-test: powershell/select_string_cmdlet/select_string_context_asymmetric
# -Context pre, post allows requesting different line counts before and after
$lines = @("first line", "second line target", "third line", "fourth line")
$match = $lines | Select-String -Pattern "target" -Context 1, 2

if ($match.Context.PreContext.Count -ne 1) {
    Write-Host "FAIL: expected 1 PreContext line, got $($match.Context.PreContext.Count)"
    exit 1
}

if ($match.Context.PostContext.Count -ne 2) {
    Write-Host "FAIL: expected 2 PostContext lines, got $($match.Context.PostContext.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
