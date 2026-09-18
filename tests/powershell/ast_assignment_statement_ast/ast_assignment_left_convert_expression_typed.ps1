# vybe-test: powershell/ast_assignment_statement_ast/ast_assignment_left_convert_expression_typed
# Type-constrained assignment [int]$count = 42 sets Left to ConvertExpressionAst
$code = "[int]`$count = 42"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assignAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)

if ($assignAst.Left -isnot [System.Management.Automation.Language.ConvertExpressionAst]) {
    Write-Host "FAIL: expected ConvertExpressionAst for type-constrained left, got '$($assignAst.Left.GetType().Name)'"
    exit 1
}

$typeConstraint = $assignAst.Left.Type
if ($typeConstraint.TypeName.FullName -ne "int") {
    Write-Host "FAIL: type constraint mismatch: '$($typeConstraint.TypeName.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
