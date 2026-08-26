# vybe-test: powershell/classes_hidden_members/hidden_static_method
class Security {
    hidden static [int]Hash([int]$x) { return $x * 31 }
    static [int]SecureVal([int]$v) { return [Security]::Hash($v) }
}
$res = [Security]::SecureVal(5)
if ($res -ne 155) {
    Write-Host "FAIL: Hidden static method failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
