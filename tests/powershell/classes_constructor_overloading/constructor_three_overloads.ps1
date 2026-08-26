# vybe-test: powershell/classes_constructor_overloading/constructor_three_overloads
class Config {
    [string]$Host
    [int]$Port
    [bool]$Ssl
    Config() { $this.Host = "localhost"; $this.Port = 80; $this.Ssl = $false }
    Config([string]$h) { $this.Host = $h; $this.Port = 80; $this.Ssl = $false }
    Config([string]$h, [int]$p, [bool]$s) { $this.Host = $h; $this.Port = $p; $this.Ssl = $s }
}
$c1 = [Config]::new()
$c2 = [Config]::new("server.io")
$c3 = [Config]::new("server.io", 443, $true)
if ($c1.Port -ne 80 -or $c2.Host -ne "server.io" -or $c3.Ssl -ne $true) {
    Write-Host "FAIL: 3-overload constructor failed"
    exit 1
}
Write-Host "PASS"
exit 0
