# vybe-test: powershell/json_array_vs_single_object/converto_json_asarray_flag_forces_json_array
$obj = [pscustomobject]@{ Id = 1 }
$json = $obj | ConvertTo-Json -AsArray
if (-not $json.Trim().StartsWith("[")) {
    Write-Host "FAIL: ConvertTo-Json -AsArray should produce JSON array brackets, got '$json'"
    exit 1
}
Write-Host "PASS"
exit 0
