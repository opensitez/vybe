# vybe-test: powershell/interpolation/method_call_quoted_arg
# The subexpression carries a method call whose ARGUMENT is a quoted string.
$s = 'abc'
$text = "r=$($s.Replace('a','X'))"
if ($text -ne 'r=Xbc') {
    Write-Host "FAIL: got [$text]"
    exit 1
}
Write-Host 'PASS'
exit 0
