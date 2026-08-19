# vybe-test: powershell/lexical_escape_rules/escape_unicode_surrogate_pair
$grin = "\u{1F600}"
if ($grin -notmatch "^[\u{1F600}]$") {
    Write-Host "FAIL: expected emoji glyph, got $grin"
    exit 1
}

Write-Host 'PASS'
exit 0
