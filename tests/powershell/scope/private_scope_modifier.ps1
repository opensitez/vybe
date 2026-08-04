# vybe-test: powershell/scope/private_scope_modifier
function Outer {
    $private:x = 99
    function Inner { return $x }   # cannot see $private:x
    return Inner
}
$result = Outer
# $private:x is not visible inside Inner, so $x should be $null / empty
if ($result -eq 99) {
    Write-Host "FAIL: private variable should not be visible in child scope"
    exit 1
}
Write-Host "PASS"
exit 0
