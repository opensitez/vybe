# vybe-test: powershell/ast_assignment_statement_ast/ast_assignment_extent_encloses_operator_and_operands
# AssignmentStatementAst.Extent encloses the left side, operator, and right side expressions
$code = "`$allocatedMemoryBytes = 1024 * 1024 * 512"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assignAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)

if ($assignAst.Extent.Text -ne $code) {
    Write-Host "FAIL: assignment extent mismatch, expected '$code', got '$($assignAst.Extent.Text)'"
    exit 1
}

if ($assignAst.Extent.StartOffset -ne 0 -or $assignAst.Extent.EndOffset -ne $code.Length) {
    Write-Host "FAIL: assignment extent offsets mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
