# vybe-test: powershell/ast_parsing/ast_parsing_array_literal_ast
$sb = { 1, 2, 3 }
$arrAst = $sb.Ast.Find({ param($ast) $ast -is [System.Management.Automation.Language.ArrayLiteralAst] }, $true)
if ($arrAst.Elements.Count -ne 3) {
    Write-Host "FAIL: ArrayLiteralAst elements expected 3, got $($arrAst.Elements.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
