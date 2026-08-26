# vybe-test: powershell/parameters_validate_script/validatescript_custom_error_message_throw
function Set-PortNumber {
    param(
        [ValidateScript({
            if ($_ -lt 1024) { throw "Port must be unprivileged (>= 1024)" }
            return $true
        })]
        [int]$Port
    )
    return $Port
}
$msg = ""
try {
    $x = Set-PortNumber -Port 80
} catch {
    $msg = $_.Exception.Message
}
if (-not $msg.Contains("Port must be unprivileged")) {
    Write-Host "FAIL: Custom exception message from ValidateScript failed, got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
