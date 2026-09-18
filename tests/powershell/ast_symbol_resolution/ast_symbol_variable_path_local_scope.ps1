# vybe-test: powershell/ast_symbol_resolution/ast_symbol_variable_path_local_scope
# Variable with $local: prefix sets IsLocal to true on VariablePath
$code = "`$local:iterationIndex = 0"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$varAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)
$vp = $varAst.VariablePath

if (-not $vp.IsLocal) {
    Write-Host "FAIL: IsLocal was false for local variable"
    exit 1
}

if ($vp.UserPath -ne "local:iterationIndex") {
    Write-Host "FAIL: UserPath mismatch: '$($vp.UserPath)'"
    exit 1
}

Write-Host "PASS"
exit 0
