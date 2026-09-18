# vybe-test: powershell/ast_symbol_resolution/ast_symbol_function_definition_name
# FunctionDefinitionAst.Name matches the declared function name
$code = "function Initialize-ApplicationEnvironment { `$true }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$funcAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)

if ($funcAst.Name -ne "Initialize-ApplicationEnvironment") {
    Write-Host "FAIL: function name mismatch: '$($funcAst.Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
