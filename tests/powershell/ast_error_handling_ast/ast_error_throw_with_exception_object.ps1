# vybe-test: powershell/ast_error_handling_ast/ast_error_throw_with_exception_object
# throw [System.InvalidOperationException]::new() parses ThrowStatementAst with constructor invocation
$code = "throw [System.InvalidOperationException]::new('Invalid config')"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$throwAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ThrowStatementAst] }, $true)

$expr = $throwAst.Pipeline.GetPureExpression()

if ($expr -isnot [System.Management.Automation.Language.InvokeMemberExpressionAst]) {
    Write-Host "FAIL: expected InvokeMemberExpressionAst on throw, got '$($expr.GetType().Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
