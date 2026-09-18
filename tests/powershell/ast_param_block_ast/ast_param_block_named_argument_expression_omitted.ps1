# vybe-test: powershell/ast_param_block_ast/ast_param_block_named_argument_expression_omitted
# NamedAttributeArgumentAst.ExpressionOmitted is true for bare switch syntax like [Parameter(Mandatory)]
$code = "param([Parameter(Mandatory)][string]`$ParamName)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$paramAst = $ast.ParamBlock.Parameters[0]
$attrAst = $paramAst.Attributes | Where-Object { $_ -is [System.Management.Automation.Language.AttributeAst] -and $_.TypeName.FullName -eq "Parameter" }
$namedArg = $attrAst.NamedArguments | Where-Object { $_.ArgumentName -eq "Mandatory" }

if ($namedArg -eq $null) {
    Write-Host "FAIL: named argument 'Mandatory' not found"
    exit 1
}

if (-not $namedArg.ExpressionOmitted) {
    Write-Host "FAIL: ExpressionOmitted was false for bare switch attribute syntax"
    exit 1
}

Write-Host "PASS"
exit 0
