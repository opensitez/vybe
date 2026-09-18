# vybe-test: powershell/ast_param_block_ast/ast_param_block_named_argument_expression_provided
# NamedAttributeArgumentAst.ExpressionOmitted is false when explicit expression is provided
$code = "param([Parameter(Mandatory = `$false, Position = 0)][string]`$ParamName)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$paramAst = $ast.ParamBlock.Parameters[0]
$attrAst = $paramAst.Attributes | Where-Object { $_ -is [System.Management.Automation.Language.AttributeAst] -and $_.TypeName.FullName -eq "Parameter" }
$mandatoryArg = $attrAst.NamedArguments | Where-Object { $_.ArgumentName -eq "Mandatory" }
$positionArg = $attrAst.NamedArguments | Where-Object { $_.ArgumentName -eq "Position" }

if ($mandatoryArg.ExpressionOmitted) {
    Write-Host "FAIL: ExpressionOmitted was true for Mandatory = `$false"
    exit 1
}

if ($positionArg.ExpressionOmitted) {
    Write-Host "FAIL: ExpressionOmitted was true for Position = 0"
    exit 1
}

if ($positionArg.Argument.Value -ne 0) {
    Write-Host "FAIL: Position argument value mismatch: '$($positionArg.Argument.Value)'"
    exit 1
}

Write-Host "PASS"
exit 0
