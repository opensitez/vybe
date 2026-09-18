# vybe-test: powershell/intrinsic_collection_overloads/foreach_parameterized_method_call
$numbers = @(10, 15, 255)

# .ForEach('MethodName', arg1, arg2) invokes the named method on each element with provided arguments
$hexStrings = $numbers.ForEach('ToString', 'X2')

if ($null -eq $hexStrings -or $hexStrings.Count -ne 3) {
    Write-Host "FAIL: .ForEach('ToString', 'X2') expected 3 items, got $($hexStrings.Count)"
    exit 1
}

if ($hexStrings[0] -ne "0A" -or $hexStrings[1] -ne "0F" -or $hexStrings[2] -ne "FF") {
    Write-Host "FAIL: expected @('0A', '0F', 'FF'), got @($($hexStrings -join ', '))"
    exit 1
}

# Also verify parameterless method call like ToUpper()
$words = @("vybe", "powershell").ForEach('ToUpper')
if ($words[0] -ne "VYBE" -or $words[1] -ne "POWERSHELL") {
    Write-Host "FAIL: .ForEach('ToUpper') expected @('VYBE', 'POWERSHELL'), got @($($words -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
