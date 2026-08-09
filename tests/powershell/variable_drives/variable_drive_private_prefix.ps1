# vybe-test: powershell/variable_drives/variable_drive_private_prefix
$private:privVar = "Hidden"
function Child-Read {
    return $privVar
}
$res = Child-Read
if ($res -ne $null) {
    Write-Host "FAIL: \$private: variable leaked to child scope"
    exit 1
}
Write-Host "PASS"
exit 0
