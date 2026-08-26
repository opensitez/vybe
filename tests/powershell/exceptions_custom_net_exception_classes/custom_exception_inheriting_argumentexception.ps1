# vybe-test: powershell/exceptions_custom_net_exception_classes/custom_exception_inheriting_argumentexception
class MissingHeaderException : System.ArgumentException {
    MissingHeaderException([string]$paramName) : base("Header is missing", $paramName) {}
}
$mhe = [MissingHeaderException]::new("Authorization")
if ($mhe.ParamName -ne "Authorization" -or $mhe -isnot [System.ArgumentException]) {
    Write-Host "FAIL: Custom exception inheriting ArgumentException failed"
    exit 1
}
Write-Host "PASS"
exit 0
