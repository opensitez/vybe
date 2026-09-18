# vybe-test: powershell/ast_param_block_ast/ast_param_block_type_constraint_ast_type_name
# ParameterAst.Attributes contains a TypeConstraintAst reflecting the type name specified in source
$code = "param([System.Guid]`$RecordId)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$paramAst = $ast.ParamBlock.Parameters[0]
$tc = $paramAst.Attributes | Where-Object { $_ -is [System.Management.Automation.Language.TypeConstraintAst] }

if ($tc -eq $null) {
    Write-Host "FAIL: TypeConstraintAst missing on parameter attributes"
    exit 1
}

if ($tc.TypeName.FullName -ne "System.Guid") {
    Write-Host "FAIL: expected TypeName 'System.Guid', got '$($tc.TypeName.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
