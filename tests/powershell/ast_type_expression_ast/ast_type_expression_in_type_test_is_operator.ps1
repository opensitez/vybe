# vybe-test: powershell/ast_type_expression_ast/ast_type_expression_in_type_test_is_operator
# In '$val -is [int]', the right operand of BinaryExpressionAst is a TypeExpressionAst
$code = "`$result = (`$val -is [int])"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$binAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.BinaryExpressionAst] }, $true)

if ($binAst.Operator -ne [System.Management.Automation.Language.TokenKind]::Is) {
    Write-Host "FAIL: binary operator was not 'Is'"
    exit 1
}

$rightNode = $binAst.Right

if ($rightNode -isnot [System.Management.Automation.Language.TypeExpressionAst]) {
    Write-Host "FAIL: expected TypeExpressionAst on right of -is operator, got '$($rightNode.GetType().Name)'"
    exit 1
}

if ($rightNode.TypeName.FullName -ne "int") {
    Write-Host "FAIL: type name mismatch: '$($rightNode.TypeName.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
