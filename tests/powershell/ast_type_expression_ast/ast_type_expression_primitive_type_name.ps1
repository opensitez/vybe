# vybe-test: powershell/ast_type_expression_ast/ast_type_expression_primitive_type_name
# [int] parses as a TypeExpressionAst with TypeName.FullName equal to 'int'
$code = "`$typeRef = [int]"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$typeAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TypeExpressionAst] }, $true)

if ($typeAst -eq $null) {
    Write-Host "FAIL: TypeExpressionAst not found"
    exit 1
}

if ($typeAst.TypeName.FullName -ne "int") {
    Write-Host "FAIL: expected FullName 'int', got '$($typeAst.TypeName.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
