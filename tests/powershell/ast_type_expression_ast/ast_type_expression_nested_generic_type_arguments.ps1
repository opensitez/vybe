# vybe-test: powershell/ast_type_expression_ast/ast_type_expression_nested_generic_type_arguments
# Nested generic types [List[List[string]]] parse as recursively generic ITypeName instances
$code = "`$nested = [System.Collections.Generic.List[System.Collections.Generic.List[string]]]"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$typeAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TypeExpressionAst] }, $true)
$outerTn = $typeAst.TypeName

if (-not $outerTn.IsGeneric) {
    Write-Host "FAIL: outer type was not generic"
    exit 1
}

$innerTn = $outerTn.GenericArguments[0]
if (-not $innerTn.IsGeneric) {
    Write-Host "FAIL: inner argument was not generic"
    exit 1
}

$leafArg = $innerTn.GenericArguments[0]
if ($leafArg.FullName -ne "string") {
    Write-Host "FAIL: leaf generic argument mismatch: '$($leafArg.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
