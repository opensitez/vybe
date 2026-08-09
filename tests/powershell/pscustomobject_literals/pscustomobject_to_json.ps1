# vybe-test: powershell/pscustomobject_literals/pscustomobject_to_json
$obj = [pscustomobject]@{ K = "V" }
$json = $obj | ConvertTo-Json -Compress
if ($json -ne '{"K":"V"}') {
    Write-Host "FAIL: ConvertTo-Json expected '`{"K`":`"V`"`}', got $json"
    exit 1
}
Write-Host "PASS"
exit 0
