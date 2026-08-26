# vybe-test: powershell/classes_hidden_members/hidden_property_read_and_write
class SecretBox {
    hidden [string]$Secret
    [string]$Label
    SecretBox([string]$s, [string]$l) {
        $this.Secret = $s
        $this.Label = $l
    }
    [string]GetSecret() { return $this.Secret }
}
$box = [SecretBox]::new("TopSecret123", "PublicLabel")
if ($box.Label -ne "PublicLabel" -or $box.GetSecret() -ne "TopSecret123" -or $box.Secret -ne "TopSecret123") {
    Write-Host "FAIL: Hidden property access failed"
    exit 1
}
Write-Host "PASS"
exit 0
