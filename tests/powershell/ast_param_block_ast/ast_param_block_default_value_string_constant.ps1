# vybe-test: powershell/ast_param_block_ast/ast_param_block_default_value_string_constant
# ParameterAst.DefaultValue parses a string default as a StringConstantExpressionAst
$code = "param([string]`$DefaultHost = 'prod-cluster-01')"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$paramAst = $ast.ParamBlock.Parameters[0]
$defVal = $paramAst.DefaultValue

if ($defVal -eq $null) {
    Write-Host "FAIL: DefaultValue was null"
    exit 1
}

if ($defVal -isnot [System.Management.Automation.Language.StringConstantExpressionAst]) {
    Write-Host "FAIL: expected StringConstantExpressionAst, got '$($defVal.GetType().Name)'"
    exit 1
}

if ($defVal.Value -ne "prod-cluster-01") {
    Write-Host "FAIL: DefaultValue mismatch, expected 'prod-cluster-01', got '$($defVal.Value)'"
    exit 1
}

Write-Host "PASS"
exit 0
