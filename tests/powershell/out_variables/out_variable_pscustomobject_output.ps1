# vybe-test: powershell/out_variables/out_variable_pscustomobject_output
[pscustomobject]@{ Tag = "OV" } | ForEach-Object { $_ } -OutVariable oCap | Out-Null
if ($oCap[0].Tag -ne "OV") {
    Write-Host "FAIL: PSCustomObject OutVariable expected Tag=OV"
    exit 1
}
Write-Host "PASS"
exit 0
