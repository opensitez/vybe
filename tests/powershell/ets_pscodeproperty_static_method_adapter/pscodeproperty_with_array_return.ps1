# vybe-test: powershell/ets_pscodeproperty_static_method_adapter/pscodeproperty_with_array_return
class ArrCode {
    static [string[]]GetTokens([psobject]$i) { return $i.Raw.Split(';') }
}
$obj = [pscustomobject]@{ Raw = "a;b;c" }
$obj.PSObject.Properties.Add([System.Management.Automation.PSCodeProperty]::new("Tokens", [ArrCode].GetMethod("GetTokens")))
if ($obj.Tokens.Length -ne 3 -or $obj.Tokens[1] -ne "b") {
    Write-Host "FAIL: PSCodeProperty with array return failed"
    exit 1
}
Write-Host "PASS"
exit 0
