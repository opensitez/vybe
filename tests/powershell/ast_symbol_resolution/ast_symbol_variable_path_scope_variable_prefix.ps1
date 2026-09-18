# vybe-test: powershell/ast_symbol_resolution/ast_symbol_variable_path_scope_variable_prefix
# Variable with $variable: prefix sets IsVariable to true while IsDriveQualified remains false
$code = "`$variable:ConfigEntry = 'Valid'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$varAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)
$vp = $varAst.VariablePath

if (-not $vp.IsVariable) {
    Write-Host "FAIL: IsVariable was false for variable: prefix"
    exit 1
}

if ($vp.IsDriveQualified) {
    Write-Host "FAIL: IsDriveQualified was unexpectedly true for variable: prefix"
    exit 1
}

if ($vp.UserPath -ne "variable:ConfigEntry") {
    Write-Host "FAIL: UserPath mismatch: '$($vp.UserPath)'"
    exit 1
}

Write-Host "PASS"
exit 0
