# vybe-test: powershell/ast_param_block_ast/ast_param_block_default_value_absent_is_null
# ParameterAst.DefaultValue is $null when no default value assignment is present
$code = "param([string]`$MandatoryKey)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$paramAst = $ast.ParamBlock.Parameters[0]

if ($paramAst.DefaultValue -ne $null) {
    Write-Host "FAIL: DefaultValue was not null when no default was provided"
    exit 1
}

Write-Host "PASS"
exit 0
