# vybe-test: powershell/ast_parsing/ast_parsing_parameter_ast
$sb = { param([string]$Name) }
$paramAst = $sb.Ast.Find({ param($ast) $ast -is [System.Management.Automation.Language.ParameterAst] }, $true)
if ($paramAst.Name.VariablePath.UserPath -ne "Name") {
    Write-Host "FAIL: ParameterAst variable path expected 'Name', got '$($paramAst.Name.VariablePath.UserPath)'"
    exit 1
}
Write-Host "PASS"
exit 0
