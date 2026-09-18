# vybe-test: powershell/comparison/case_sensitive_clike_distinguishes_case
# The -clike operator requires exact letter-case agreement when matching wildcards
$value = "AlphaBravo"

$matchesProperCase = $value -clike "Alpha*"
$matchesWrongCase = $value -clike "alpha*"

if (-not $matchesProperCase) {
    Write-Host "FAIL: expected -clike to match identical case"
    exit 1
}

if ($matchesWrongCase) {
    Write-Host "FAIL: -clike should not match differing letter case"
    exit 1
}

Write-Host "PASS"
exit 0
