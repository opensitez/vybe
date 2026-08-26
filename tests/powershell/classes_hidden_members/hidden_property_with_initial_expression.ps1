# vybe-test: powershell/classes_hidden_members/hidden_property_with_initial_expression
class AutoInit {
    hidden [datetime]$Created = [datetime]::UtcNow
    [datetime]GetCreated() { return $this.Created }
}
$ai = [AutoInit]::new()
if ($ai.GetCreated().Year -lt 2026) {
    Write-Host "FAIL: AutoInit hidden property failed"
    exit 1
}
Write-Host "PASS"
exit 0
