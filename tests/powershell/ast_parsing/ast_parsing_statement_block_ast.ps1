# vybe-test: powershell/ast_parsing/ast_parsing_statement_block_ast
$sb = { $x = 1 }
$block = $sb.Ast.EndBlock
if ($block -eq $null -or $block.Statements.Count -ne 1) {
    Write-Host "FAIL: EndBlock statements count expected 1"
    exit 1
}
Write-Host "PASS"
exit 0
