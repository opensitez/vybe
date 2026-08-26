# vybe-test: powershell/string_builder_operations/maxcapacity_enforcement
$sb = [System.Text.StringBuilder]::new(2, 4)
$null = $sb.Append("1234")
$caught = $false
try {
    $null = $sb.Append("5")
} catch [System.ArgumentOutOfRangeException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected ArgumentOutOfRangeException on MaxCapacity overflow"
    exit 1
}
Write-Host "PASS"
exit 0
