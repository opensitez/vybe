# vybe-test: powershell/ast_pipeline_ast_inspection/ast_pipeline_invocation_operator_ampersand
# & command invocation sets CommandAst.InvocationOperator to TokenKind.Ampersand
$code = "& 'git' status"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$cmdAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)

if ($cmdAst.InvocationOperator -ne [System.Management.Automation.Language.TokenKind]::Ampersand) {
    Write-Host "FAIL: InvocationOperator was not Ampersand, got '$($cmdAst.InvocationOperator)'"
    exit 1
}

Write-Host "PASS"
exit 0
