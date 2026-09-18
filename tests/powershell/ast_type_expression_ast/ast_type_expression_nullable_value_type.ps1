# vybe-test: powershell/ast_type_expression_ast/ast_type_expression_nullable_value_type
# [System.Nullable[int]] parses as a generic type and resolves to Nullable`1[Int32]
$code = "`$nullableInt = [System.Nullable[int]]"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$typeAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TypeExpressionAst] }, $true)
$refType = $typeAst.TypeName.GetReflectionType()

if ($refType -ne [System.Nullable[int]]) {
    Write-Host "FAIL: nullable reflection type mismatch: '$($refType.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
