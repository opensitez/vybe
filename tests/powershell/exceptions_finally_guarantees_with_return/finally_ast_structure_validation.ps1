# vybe-test: powershell/exceptions_finally_guarantees_with_return/finally_ast_structure_validation
$ast = [System.Management.Automation.Language.Parser]::ParseInput('
try { 1 } finally { 2 }
', [ref]$null, [ref]$null)
$tryAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TryStatementAst] }, $true)
if ($tryAst.Finally -eq $null) {
    Write-Host "FAIL: TryStatementAst Finally AST validation failed"
    exit 1
}
Write-Host "PASS"
exit 0
