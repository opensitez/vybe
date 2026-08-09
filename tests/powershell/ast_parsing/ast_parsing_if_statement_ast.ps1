# vybe-test: powershell/ast_parsing/ast_parsing_if_statement_ast
$sb = { if ($true) { "YES" } else { "NO" } }
$ifAst = $sb.Ast.Find({ param($ast) $ast -is [System.Management.Automation.Language.IfStatementAst] }, $true)
if ($ifAst.Clauses.Count -ne 1 -or $ifAst.ElseClause -eq $null) {
    Write-Host "FAIL: IfStatementAst clauses/else clause missing"
    exit 1
}
Write-Host "PASS"
exit 0
