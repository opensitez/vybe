# vybe-test: powershell/out_variables/out_variable_type_array
"Single" | ForEach-Object { $_ } -OutVariable arr | Out-Null
if (-not ($arr -is [System.Collections.ArrayList])) {
    Write-Host "FAIL: OutVariable type expected ArrayList, got $($arr.GetType().FullName)"
    exit 1
}
Write-Host "PASS"
exit 0
