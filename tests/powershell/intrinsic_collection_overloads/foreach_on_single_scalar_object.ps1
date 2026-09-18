# vybe-test: powershell/intrinsic_collection_overloads/foreach_on_single_scalar_object
# In PowerShell v3+, non-collection scalar values have intrinsic .ForEach() methods added automatically
$number = 42

$doubled = $number.ForEach({ $_ * 2 })

if ($null -eq $doubled) {
    Write-Host "FAIL: (42).ForEach() returned `$null"
    exit 1
}

# The result is wrapped in a collection of count 1
if ($doubled[0] -ne 84) {
    Write-Host "FAIL: expected 84 from scalar .ForEach(), got $($doubled[0])"
    exit 1
}

# Also verify scalar member extraction via .ForEach('Length') on a string
$str = "hello"
$len = $str.ForEach('Length')
if ($len[0] -ne 5) {
    Write-Host "FAIL: expected string length 5 from scalar .ForEach('Length'), got $($len[0])"
    exit 1
}

Write-Host "PASS"
exit 0
