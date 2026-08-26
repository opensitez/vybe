# vybe-test: powershell/parameters_alias_attribute/alias_with_pipeline_binding
function Filter-HostName {
    param(
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Alias("CN", "MachineName")]
        [string]$Name
    )
    process { "Name:$Name" }
}
$obj = [pscustomobject]@{ CN = "web-prod-01" }
$res = $obj | Filter-HostName
if ($res -ne "Name:web-prod-01") {
    Write-Host "FAIL: Pipeline by property name with alias failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
