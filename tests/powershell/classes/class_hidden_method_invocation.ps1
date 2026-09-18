# vybe-test: powershell/classes/class_hidden_method_invocation
# A method marked hidden is invokable on an instance and from within public class methods
class SecuredEndpoint {
    hidden [string] GetInternalToken() {
        return "SEC-KEY-999"
    }

    [string] Authorize() {
        return "AUTHORIZED:" + $this.GetInternalToken()
    }
}

$ep = [SecuredEndpoint]::new()

# Public method calling hidden method
$authMsg = $ep.Authorize()
if ($authMsg -ne "AUTHORIZED:SEC-KEY-999") {
    Write-Host "FAIL: expected 'AUTHORIZED:SEC-KEY-999', got '$authMsg'"
    exit 1
}

# Direct invocation of hidden method is still supported
$directToken = $ep.GetInternalToken()
if ($directToken -ne "SEC-KEY-999") {
    Write-Host "FAIL: direct hidden call expected 'SEC-KEY-999', got '$directToken'"
    exit 1
}

Write-Host "PASS"
exit 0
