# vybe-test: powershell/ast_assignment_statement_ast/ast_assignment_left_member_expression
# When assigning to an object property, Left is a MemberExpressionAst
$code = "`$user.Department = 'Infrastructure'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assignAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)

if ($assignAst.Left -isnot [System.Management.Automation.Language.MemberExpressionAst]) {
    Write-Host "FAIL: Left was not MemberExpressionAst"
    exit 1
}

$member = $assignAst.Left.Member.Value
if ($member -ne "Department") {
    Write-Host "FAIL: property name mismatch: '$member'"
    exit 1
}

Write-Host "PASS"
exit 0
