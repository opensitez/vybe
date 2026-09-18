# vybe-test: powershell/select_string_cmdlet/select_string_line_and_linenumber_properties
# MatchInfo exposes .Line with the matching text and .LineNumber with 1-indexed position
$lines = @("first header", "second record containing target", "third footer")
$match = $lines | Select-String -Pattern "target"

if ($match.Line -ne "second record containing target") {
    Write-Host "FAIL: unexpected .Line content: '$($match.Line)'"
    exit 1
}

if ($match.LineNumber -ne 2) {
    Write-Host "FAIL: expected .LineNumber 2, got: $($match.LineNumber)"
    exit 1
}

Write-Host "PASS"
exit 0
