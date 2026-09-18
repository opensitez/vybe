# vybe-test: powershell/ast_param_block_ast/ast_param_block_static_type_switch
# ParameterAst.StaticType resolves [switch] to System.Management.Automation.SwitchParameter
$code = "param([switch]`$ForceExecution)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$paramAst = $ast.ParamBlock.Parameters[0]
$staticType = $paramAst.StaticType

if ($staticType -ne [System.Management.Automation.SwitchParameter]) {
    Write-Host "FAIL: expected SwitchParameter type, got '$($staticType.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
