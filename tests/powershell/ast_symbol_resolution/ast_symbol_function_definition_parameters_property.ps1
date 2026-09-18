# vybe-test: powershell/ast_symbol_resolution/ast_symbol_function_definition_parameters_property
# FunctionDefinitionAst.Parameters contains declared parameters when defined in function header
$code = "function Multiply-Values(`$Alpha, `$Beta) { `$Alpha * `$Beta }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$funcAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)

if ($funcAst.Parameters.Count -ne 2) {
    Write-Host "FAIL: expected 2 parameters on FunctionDefinitionAst, got $($funcAst.Parameters.Count)"
    exit 1
}

$firstParamName = $funcAst.Parameters[0].Name.VariablePath.UserPath
$secondParamName = $funcAst.Parameters[1].Name.VariablePath.UserPath

if ($firstParamName -ne "Alpha" -or $secondParamName -ne "Beta") {
    Write-Host "FAIL: function parameter names mismatch: '$firstParamName', '$secondParamName'"
    exit 1
}

Write-Host "PASS"
exit 0
