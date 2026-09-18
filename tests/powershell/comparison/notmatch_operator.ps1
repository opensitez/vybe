# vybe-test: powershell/comparison/notmatch_operator
# The -notmatch operator returns true when a string does not match the regular expression
$inputString = "ErrorCode_500"

$nonNumericMatch = $inputString -notmatch "^\d+$"
$hasNumericMatch = $inputString -notmatch "\d+"

if (-not $nonNumericMatch) {
    Write-Host "FAIL: expected -notmatch to return true for non-digits-only string"
    exit 1
}

if ($hasNumericMatch) {
    Write-Host "FAIL: expected -notmatch to return false because string contains digits"
    exit 1
}

Write-Host "PASS"
exit 0
