# vybe-test: powershell/lexical_escape_rules/escape_in_regex_like
$match = 'a+b' -match 'a\+b'
if (-not $match) {
    Write-Host 'FAIL: regex backslash-plus should be treated as literal plus'
    exit 1
}

$match2 = 'ab' -match 'a\+b'
if ($match2) {
    Write-Host 'FAIL: regex special escaping not respected'
    exit 1
}

Write-Host 'PASS'
exit 0
