# vybe-test: powershell/exceptions_custom_net_exception_classes/custom_exception_with_extra_properties
class HttpError : System.Exception {
    [int]$StatusCode
    HttpError([int]$code, [string]$msg) : base($msg) {
        $this.StatusCode = $code
    }
}
$he = [HttpError]::new(404, "Page Not Found")
if ($he.StatusCode -ne 404 -or $he.Message -ne "Page Not Found") {
    Write-Host "FAIL: Custom exception with extra properties failed"
    exit 1
}
Write-Host "PASS"
exit 0
