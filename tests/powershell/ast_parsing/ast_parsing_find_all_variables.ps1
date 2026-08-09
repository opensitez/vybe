# vybe-test: powershell/ast_parsing/ast_parsing_find_all_variables
$sb = { $a = 1; $b = 2; $c = $a + $b }
$vars = $sb.Ast.FindAll({ param($ast) $ast -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)
if ($vars.Count -lt 4) {
    Write-Host "FAIL: VariableExpressionAst FindAll count expected >= 4, got $($vars.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
