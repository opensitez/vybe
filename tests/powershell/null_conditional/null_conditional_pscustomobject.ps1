# vybe-test: powershell/null_conditional/null_conditional_pscustomobject
$data = [pscustomobject]@{ Sub = $null }
$res = ${data}?.Sub?.Val
if ($res -ne $null) {
    Write-Host "FAIL: nested null PSCustomObject property expected null, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
