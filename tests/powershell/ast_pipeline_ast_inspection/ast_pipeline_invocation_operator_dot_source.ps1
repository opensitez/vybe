# vybe-test: powershell/ast_pipeline_ast_inspection/ast_pipeline_invocation_operator_dot_source
# . script invocation sets CommandAst.InvocationOperator to TokenKind.Dot
$code = ". ./initialize.ps1"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$cmdAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)

if ($cmdAst.InvocationOperator -ne [System.Management.Automation.Language.TokenKind]::Dot) {
    Write-Host "FAIL: InvocationOperator was not Dot, got '$($cmdAst.InvocationOperator)'"
    exit 1
}

Write-Host "PASS"
exit 0
