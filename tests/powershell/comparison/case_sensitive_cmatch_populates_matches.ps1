# vybe-test: powershell/comparison/case_sensitive_cmatch_populates_matches
# The -cmatch operator performs case-sensitive regex matching and populates the $Matches automatic variable
$uppercaseString = "ProjectVybe_2026"

$caseMatch = $uppercaseString -cmatch "^[A-Z][a-z]+[A-Z][a-z]+_(\d+)$"

if (-not $caseMatch) {
    Write-Host "FAIL: -cmatch failed on properly cased string"
    exit 1
}

if ($Matches[1] -ne "2026") {
    Write-Host "FAIL: expected capture group '2026', got '$($Matches[1])'"
    exit 1
}

$lowercaseString = "projectvybe_2026"
$caseMismatch = $lowercaseString -cmatch "^[A-Z]"

if ($caseMismatch) {
    Write-Host "FAIL: -cmatch unexpectedly matched lowercase string against uppercase pattern"
    exit 1
}

Write-Host "PASS"
exit 0
