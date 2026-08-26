# vybe-test: powershell/parameters_alias_attribute/alias_on_mandatory_parameter
function Set-MandatoryAlias {
    param(
        [Parameter(Mandatory=$true)]
        [Alias("Target")]
        [string]$Path
    )
    return "Path:$Path"
}
$res = Set-MandatoryAlias -Target "/var/log"
if ($res -ne "Path:/var/log") {
    Write-Host "FAIL: Alias on mandatory parameter failed"
    exit 1
}
Write-Host "PASS"
exit 0
