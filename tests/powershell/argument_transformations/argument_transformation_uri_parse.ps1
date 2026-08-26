# vybe-test: powershell/argument_transformations/argument_transformation_uri_parse
class UriTransform : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$e, [object]$i) {
        if ($i -is [string]) { return [uri]$i }
        return $i
    }
}
function Test-Uri {
    param([UriTransform()][uri]$Endpoint)
    return $Endpoint.Host
}
$res = Test-Uri "https://vybe.dev"
if ($res -ne "vybe.dev") {
    Write-Host "FAIL: UriTransform expected Host 'vybe.dev', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
