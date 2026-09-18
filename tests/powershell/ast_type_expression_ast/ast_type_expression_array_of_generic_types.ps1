# vybe-test: powershell/ast_type_expression_ast/ast_type_expression_array_of_generic_types
# [System.Collections.Generic.List[string][]] parses as an ArrayTypeName whose ElementType is generic
$code = "`$arrOfLists = [System.Collections.Generic.List[string][]]"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$typeAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TypeExpressionAst] }, $true)
$tn = $typeAst.TypeName

if (-not $tn.IsArray) {
    Write-Host "FAIL: IsArray was false for array of generic lists"
    exit 1
}

if (-not $tn.ElementType.IsGeneric) {
    Write-Host "FAIL: ElementType was not generic"
    exit 1
}

Write-Host "PASS"
exit 0
