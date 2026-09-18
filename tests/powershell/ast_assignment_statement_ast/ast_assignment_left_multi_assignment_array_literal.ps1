# vybe-test: powershell/ast_assignment_statement_ast/ast_assignment_left_multi_assignment_array_literal
# Multi-variable assignment $a, $b = 1, 2 sets Left to ArrayLiteralAst
$code = "`$first, `$second = 10, 20"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assignAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)

if ($assignAst.Left -isnot [System.Management.Automation.Language.ArrayLiteralAst]) {
    Write-Host "FAIL: Left was not ArrayLiteralAst for multi-assignment"
    exit 1
}

$elements = $assignAst.Left.Elements

if ($elements.Count -ne 2) {
    Write-Host "FAIL: expected 2 elements in Left array literal, got $($elements.Count)"
    exit 1
}

if ($elements[0].VariablePath.UserPath -ne "first" -or $elements[1].VariablePath.UserPath -ne "second") {
    Write-Host "FAIL: elements names mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
