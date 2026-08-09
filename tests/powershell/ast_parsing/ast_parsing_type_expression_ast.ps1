# vybe-test: powershell/ast_parsing/ast_parsing_type_expression_ast
$sb = { [int]"42" }
$typeAst = $sb.Ast.Find({ param($ast) $ast -is [System.Management.Automation.Language.TypeExpressionAst] }, $true)
if ($typeAst.TypeName.Name -ne "int") {
    Write-Host "FAIL: TypeExpressionAst TypeName expected 'int', got '$($typeAst.TypeName.Name)'"
    exit 1
}
Write-Host "PASS"
exit 0
