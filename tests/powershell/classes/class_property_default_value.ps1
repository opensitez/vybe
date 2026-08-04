# vybe-test: powershell/classes/class_property_default_value
class Config {
    [string]$Host     = "localhost"
    [int]$Port        = 8080
    [bool]$Secure     = $false
}
$cfg = [Config]::new()
if ($cfg.Host -ne "localhost") { Write-Host "FAIL: Host"; exit 1 }
if ($cfg.Port -ne 8080)        { Write-Host "FAIL: Port"; exit 1 }
if ($cfg.Secure -ne $false)    { Write-Host "FAIL: Secure"; exit 1 }
Write-Host "PASS"
exit 0
