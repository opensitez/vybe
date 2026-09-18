# vybe-test: powershell/ast_param_block_ast/ast_param_block_static_type_untyped_defaults_to_object
# An untyped parameter defaults StaticType to System.Object
$code = "param(`$UntypedArgument)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$paramAst = $ast.ParamBlock.Parameters[0]
$staticType = $paramAst.StaticType

if ($staticType -ne [object]) {
    Write-Host "FAIL: expected StaticType [object], got '$($staticType.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
