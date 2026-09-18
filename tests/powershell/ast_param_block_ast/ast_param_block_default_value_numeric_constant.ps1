# vybe-test: powershell/ast_param_block_ast/ast_param_block_default_value_numeric_constant
# ParameterAst.DefaultValue parses a numeric constant default as ConstantExpressionAst
$code = "param([int]`$MaxRetries = 5)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$paramAst = $ast.ParamBlock.Parameters[0]
$defVal = $paramAst.DefaultValue

if ($defVal -eq $null) {
    Write-Host "FAIL: DefaultValue was null"
    exit 1
}

if ($defVal -isnot [System.Management.Automation.Language.ConstantExpressionAst]) {
    Write-Host "FAIL: expected ConstantExpressionAst, got '$($defVal.GetType().Name)'"
    exit 1
}

if ($defVal.Value -ne 5) {
    Write-Host "FAIL: DefaultValue mismatch, expected 5, got '$($defVal.Value)'"
    exit 1
}

Write-Host "PASS"
exit 0
