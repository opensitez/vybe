# vybe-test: powershell/ast_type_expression_ast/ast_type_expression_generic_argument_reflection_types
# Calling GetReflectionType on a generic TypeName returns the concrete closed generic Type
$code = "`$ref = [System.Collections.Generic.List[int]]"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$typeAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TypeExpressionAst] }, $true)
$concreteType = $typeAst.TypeName.GetReflectionType()

if ($concreteType -ne [System.Collections.Generic.List[int]]) {
    Write-Host "FAIL: reflection type mismatch, expected List[int], got '$($concreteType.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
