# vybe-test: powershell/ast_type_expression_ast/ast_type_expression_in_type_conversion_as_operator
# In '$val -as [string]', the right operand of BinaryExpressionAst is a TypeExpressionAst
$code = "`$converted = (`$val -as [string])"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$binAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.BinaryExpressionAst] }, $true)

if ($binAst.Operator -ne [System.Management.Automation.Language.TokenKind]::As) {
    Write-Host "FAIL: binary operator was not 'As'"
    exit 1
}

$rightNode = $binAst.Right

if ($rightNode -isnot [System.Management.Automation.Language.TypeExpressionAst]) {
    Write-Host "FAIL: expected TypeExpressionAst on right of -as operator, got '$($rightNode.GetType().Name)'"
    exit 1
}

if ($rightNode.TypeName.FullName -ne "string") {
    Write-Host "FAIL: type name mismatch: '$($rightNode.TypeName.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
