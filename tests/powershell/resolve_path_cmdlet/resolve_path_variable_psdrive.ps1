# vybe-test: powershell/resolve_path_cmdlet/resolve_path_variable_psdrive
# Resolve-Path against a Variable PSDrive path resolves to the 'Variable' provider
$info = Resolve-Path "Variable:PWD"

if ($null -eq $info) {
    Write-Host "FAIL: Resolve-Path returned `$null for Variable:PWD"
    exit 1
}

if ($info.Provider.Name -ne "Variable") {
    Write-Host "FAIL: expected provider 'Variable', got: '$($info.Provider.Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
