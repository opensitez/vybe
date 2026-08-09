# vybe-test: powershell/null_conditional/null_conditional_array_of_objects
$items = @([pscustomobject]@{ Id = 1 }, $null, [pscustomobject]@{ Id = 3 })
$res = $items | ForEach-Object { ${_}?.Id }
if ($res[0] -ne 1 -or $res[1] -ne $null -or $res[2] -ne 3) {
    Write-Host "FAIL: array null-conditional property expected 1, null, 3"
    exit 1
}
Write-Host "PASS"
exit 0
