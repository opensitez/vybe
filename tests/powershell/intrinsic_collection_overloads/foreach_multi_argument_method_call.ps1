# vybe-test: powershell/intrinsic_collection_overloads/foreach_multi_argument_method_call
$desserts = @("apple-pie", "banana-pie", "cherry-pie")

# .ForEach('MethodName', arg1, arg2) passes multiple arguments to the method instance on each element
$cakes = $desserts.ForEach('Replace', 'pie', 'cake')

if ($null -eq $cakes -or $cakes.Count -ne 3) {
    Write-Host "FAIL: .ForEach('Replace', ...) expected 3 items, got $($cakes.Count)"
    exit 1
}

if ($cakes[0] -ne "apple-cake" -or $cakes[1] -ne "banana-cake" -or $cakes[2] -ne "cherry-cake") {
    Write-Host "FAIL: expected cake substitutions, got @($($cakes -join ', '))"
    exit 1
}

# Also test another multi-argument method: String.Substring(startIndex, length)
$sub = @("abcdef", "uvwxyz").ForEach('Substring', 1, 3)
if ($sub[0] -ne "bcd" -or $sub[1] -ne "vwx") {
    Write-Host "FAIL: .ForEach('Substring', 1, 3) expected @('bcd', 'vwx'), got @($($sub -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
