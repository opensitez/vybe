# vybe-test: powershell/ast_param_block_ast/ast_param_block_static_type_array_constraint
# ParameterAst.StaticType accurately resolves array type constraints like [string[]]
$code = "param([string[]]`$PathList)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$paramAst = $ast.ParamBlock.Parameters[0]
$staticType = $paramAst.StaticType

if ($staticType -ne [string[]]) {
    Write-Host "FAIL: expected StaticType [string[]], got '$($staticType.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
