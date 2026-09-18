# vybe-test: powershell/comparison/is_and_isnot_type_comparison_operators
# The -is and -isnot operators perform type-membership checks against .NET type literals
$integerVal = 1024
$stringVal = "VybePlatform"

if (-not ($integerVal -is [int])) {
    Write-Host "FAIL: expected integerVal -is [int]"
    exit 1
}

if (-not ($integerVal -isnot [string])) {
    Write-Host "FAIL: expected integerVal -isnot [string]"
    exit 1
}

if (-not ($stringVal -is [string])) {
    Write-Host "FAIL: expected stringVal -is [string]"
    exit 1
}

if ($stringVal -is [int]) {
    Write-Host "FAIL: unexpected stringVal -is [int]"
    exit 1
}

Write-Host "PASS"
exit 0
