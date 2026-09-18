# vybe-test: powershell/ast_assignment_statement_ast/ast_assignment_error_position_extent
# AssignmentStatementAst.ErrorPosition returns an accurate IScriptExtent centered on the assignment operator
$code = "`$target = 'assigned value'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assignAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)

$errPos = $assignAst.ErrorPosition

if ($errPos -eq $null) {
    Write-Host "FAIL: ErrorPosition was null"
    exit 1
}

if ($errPos.Text -ne "=") {
    Write-Host "FAIL: ErrorPosition text was not '=', got '$($errPos.Text)'"
    exit 1
}

Write-Host "PASS"
exit 0
