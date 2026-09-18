# vybe-test: powershell/ast_symbol_resolution/ast_symbol_variable_path_global_scope
# Variable with $global: prefix sets IsGlobal to true on VariablePath
$code = "`$global:SharedConfig = 'ClusterA'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$varAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)
$vp = $varAst.VariablePath

if (-not $vp.IsGlobal) {
    Write-Host "FAIL: IsGlobal was false for global variable"
    exit 1
}

if ($vp.UserPath -ne "global:SharedConfig") {
    Write-Host "FAIL: UserPath mismatch: '$($vp.UserPath)'"
    exit 1
}

Write-Host "PASS"
exit 0
