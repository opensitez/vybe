# vybe-test: powershell/ast_param_block_ast/ast_param_block_default_value_subexpression
# ParameterAst.DefaultValue parses $(...) default assignments as SubExpressionAst
$code = "param(`$Timestamp = `$(Get-Date))"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$paramAst = $ast.ParamBlock.Parameters[0]
$defVal = $paramAst.DefaultValue

if ($defVal -isnot [System.Management.Automation.Language.SubExpressionAst]) {
    Write-Host "FAIL: expected SubExpressionAst for `$Timestamp default, got '$($defVal.GetType().Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
