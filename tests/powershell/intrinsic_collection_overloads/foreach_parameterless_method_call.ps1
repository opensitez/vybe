# vybe-test: powershell/intrinsic_collection_overloads/foreach_parameterless_method_call
$strings = @("alpha", "beta", "gamma")

# .ForEach('MethodName') without arguments invokes the parameterless instance method on each element
$upperStrings = $strings.ForEach('ToUpper')

if ($null -eq $upperStrings -or $upperStrings.Count -ne 3) {
    Write-Host "FAIL: .ForEach('ToUpper') expected 3 items, got $($upperStrings.Count)"
    exit 1
}

if ($upperStrings[0] -ne "ALPHA" -or $upperStrings[1] -ne "BETA" -or $upperStrings[2] -ne "GAMMA") {
    Write-Host "FAIL: expected @('ALPHA', 'BETA', 'GAMMA'), got @($($upperStrings -join ', '))"
    exit 1
}

# Also test another parameterless method: String.Trim()
$spaced = @("  one  ", "  two  ").ForEach('Trim')
if ($spaced[0] -ne "one" -or $spaced[1] -ne "two") {
    Write-Host "FAIL: .ForEach('Trim') expected @('one', 'two'), got @($($spaced -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
