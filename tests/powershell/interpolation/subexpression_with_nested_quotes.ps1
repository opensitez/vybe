# vybe-test: powershell/interpolation/subexpression_with_nested_quotes
# Subexpressions inside double quotes can contain nested strings with identical quotes without conflict
$serviceName = "authentication"

$banner = "Service status: $("active: " + $serviceName.ToUpper())"

$expected = "Service status: active: AUTHENTICATION"
if ($banner -ne $expected) {
    Write-Host "FAIL: expected '$expected', got '$banner'"
    exit 1
}

Write-Host "PASS"
exit 0
