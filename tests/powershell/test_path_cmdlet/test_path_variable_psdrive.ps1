# vybe-test: powershell/test_path_cmdlet/test_path_variable_psdrive
# Test-Path checks existence of PowerShell variables via the Variable: PSDrive
$exists = Test-Path "Variable:PWD"
$missing = Test-Path "Variable:CompletelyUndefinedVariableXYZ99"

if ($exists -ne $true) {
    Write-Host "FAIL: Test-Path 'Variable:PWD' returned `$false"
    exit 1
}

if ($missing -ne $false) {
    Write-Host "FAIL: Test-Path on missing variable returned `$true"
    exit 1
}

Write-Host "PASS"
exit 0
