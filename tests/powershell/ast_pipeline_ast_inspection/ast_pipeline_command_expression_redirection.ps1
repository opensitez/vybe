# vybe-test: powershell/ast_pipeline_ast_inspection/ast_pipeline_command_expression_redirection
# Output redirection on parenthesized expressions attaches FileRedirectionAst to CommandExpressionAst
$code = "(1..100) > 'numbers.txt'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$cmdExpr = $ast.Find({ $args[0] -is [System.Management.Automation.Language.CommandExpressionAst] }, $true)

if ($cmdExpr.Redirections.Count -ne 1) {
    Write-Host "FAIL: expected 1 redirection on CommandExpressionAst, got $($cmdExpr.Redirections.Count)"
    exit 1
}

$redir = $cmdExpr.Redirections[0]
if ($redir -isnot [System.Management.Automation.Language.FileRedirectionAst]) {
    Write-Host "FAIL: expected FileRedirectionAst on CommandExpressionAst"
    exit 1
}

Write-Host "PASS"
exit 0
