# vybe-test: powershell/ast_pipeline_ast_inspection/ast_pipeline_invocation_operator_default_unknown
# Standard direct command invocation sets CommandAst.InvocationOperator to TokenKind.Unknown
$code = "Get-ChildItem -Path 'C:\'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$cmdAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)

if ($cmdAst.InvocationOperator -ne [System.Management.Automation.Language.TokenKind]::Unknown) {
    Write-Host "FAIL: standard invocation operator was not Unknown, got '$($cmdAst.InvocationOperator)'"
    exit 1
}

Write-Host "PASS"
exit 0
