# vybe-test: powershell/pscmdlet_transaction_binding/transaction_binding_ast_expression_boolean_value
# The argument value of SupportsTransactions in the AST resolves to a VariableExpressionAst representing $true
$code = '
function CheckAstBooleanArg {
    [CmdletBinding(SupportsTransactions = $true)]
    param()
}
'

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$funcAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
$attrAst = $funcAst.Body.ParamBlock.Attributes[0]
$namedArg = $attrAst.NamedArguments | Where-Object { $_.ArgumentName -eq "SupportsTransactions" }

$argExpr = $namedArg.Argument
if ($argExpr -isnot [System.Management.Automation.Language.VariableExpressionAst]) {
    Write-Host "FAIL: expected VariableExpressionAst, got $($argExpr.GetType().Name)"
    exit 1
}

if ($argExpr.VariablePath.UserPath -ne "true") {
    Write-Host "FAIL: expected variable path 'true', got '$($argExpr.VariablePath.UserPath)'"
    exit 1
}

Write-Host "PASS"
exit 0
