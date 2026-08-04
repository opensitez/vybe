# vybe-test: powershell/variables/constant_variable
Set-Variable -Name PI -Value 3.14159 -Option Constant
$result = [Math]::Round($PI * 2, 5)
if ($result -ne 6.28318) {
    Write-Host "FAIL: expected 6.28318, got $result"
    exit 1
}
# Attempting to change a constant should throw
$threw = $false
try { Set-Variable -Name PI -Value 0 } catch { $threw = $true }
if (-not $threw) { Write-Host "FAIL: should have thrown on constant reassign"; exit 1 }
Write-Host "PASS"
exit 0
