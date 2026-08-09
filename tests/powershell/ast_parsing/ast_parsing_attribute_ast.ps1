# vybe-test: powershell/ast_parsing/ast_parsing_attribute_ast
$sb = { param([Parameter(Mandatory=$true)][string]$Text) }
$attrAst = $sb.Ast.Find({ param($ast) $ast -is [System.Management.Automation.Language.AttributeAst] }, $true)
if ($attrAst.TypeName.Name -ne "Parameter") {
    Write-Host "FAIL: AttributeAst TypeName expected 'Parameter', got '$($attrAst.TypeName.Name)'"
    exit 1
}
Write-Host "PASS"
exit 0
