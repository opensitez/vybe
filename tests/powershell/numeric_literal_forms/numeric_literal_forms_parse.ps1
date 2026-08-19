# vybe-test: powershell/numeric_literal_forms/parse
$errors = @()
[void][System.Management.Automation.PSParser]::Tokenize('12 + 34', [ref]$errors)
if ($errors.Count -ne 0) {
  Write-Host "FAIL: simple parse should produce no errors"
  exit 1
}
Write-Host 'PASS'
exit 0
