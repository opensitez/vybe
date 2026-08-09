# vybe-test: powershell/ast_parsing/ast_parsing_scriptblock_ast
$sb = { { Write-Host "Nested" } }
$sbAst = $sb.Ast.Find({ param($ast) $ast -is [System.Management.Automation.Language.ScriptBlockExpressionAst] }, $true)
if ($sbAst -eq $null) {
    Write-Host "FAIL: ScriptBlockExpressionAst missing from AST"
    exit 1
}
Write-Host "PASS"
exit 0
