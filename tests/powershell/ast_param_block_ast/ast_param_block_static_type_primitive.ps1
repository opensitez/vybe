# vybe-test: powershell/ast_param_block_ast/ast_param_block_static_type_primitive
# ParameterAst.StaticType accurately resolves primitive numeric types like [int] to System.Int32
$code = "param([int]`$TimeoutSeconds)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$paramAst = $ast.ParamBlock.Parameters[0]
$staticType = $paramAst.StaticType

if ($staticType -ne [int]) {
    Write-Host "FAIL: expected StaticType [int], got '$($staticType.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
