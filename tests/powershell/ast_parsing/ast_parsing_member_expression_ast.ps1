# vybe-test: powershell/ast_parsing/ast_parsing_member_expression_ast
$sb = { $obj.Property }
$memAst = $sb.Ast.Find({ param($ast) $ast -is [System.Management.Automation.Language.MemberExpressionAst] }, $true)
if ($memAst.Member.Value -ne "Property") {
    Write-Host "FAIL: MemberExpressionAst Member value expected 'Property', got '$($memAst.Member.Value)'"
    exit 1
}
Write-Host "PASS"
exit 0
