# vybe-test: powershell/error_record_exception_chaining/chained_exception_with_custom_class_exceptions
class InnerError : System.Exception {
    InnerError([string]$m) : base($m) {}
}
class OuterError : System.Exception {
    OuterError([string]$m, [System.Exception]$inner) : base($m, $inner) {}
}
$in = [InnerError]::new("InnerMsg")
$out = [OuterError]::new("OuterMsg", $in)
if ($out.InnerException -isnot [InnerError] -or $out.InnerException.Message -ne "InnerMsg") {
    Write-Host "FAIL: Custom class exception chaining failed"
    exit 1
}
Write-Host "PASS"
exit 0
