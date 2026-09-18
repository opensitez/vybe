# vybe-test: powershell/ast_param_block_ast/ast_param_block_parameter_name_resolution
# ParameterAst.Name.VariablePath.UserPath returns the parameter name stripped of the dollar sign
$code = "param(`$ClusterEndpoint)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$paramAst = $ast.ParamBlock.Parameters[0]
$name = $paramAst.Name.VariablePath.UserPath

if ($name -ne "ClusterEndpoint") {
    Write-Host "FAIL: expected parameter name 'ClusterEndpoint', got '$name'"
    exit 1
}

Write-Host "PASS"
exit 0
