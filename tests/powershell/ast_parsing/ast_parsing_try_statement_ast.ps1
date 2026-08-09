# vybe-test: powershell/ast_parsing/ast_parsing_try_statement_ast
$sb = { try { 1 } catch { 2 } finally { 3 } }
$tryAst = $sb.Ast.Find({ param($ast) $ast -is [System.Management.Automation.Language.TryStatementAst] }, $true)
if ($tryAst.CatchClauses.Count -ne 1 -or $tryAst.Finally -eq $null) {
    Write-Host "FAIL: TryStatementAst CatchClauses/Finally block missing"
    exit 1
}
Write-Host "PASS"
exit 0
