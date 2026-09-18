# vybe-test: powershell/ast_type_expression_ast/ast_type_expression_enum_type_resolution
# Enum types like [System.IO.FileAccess] parse as TypeExpressionAst and resolve to Enum reflection type
$code = "`$accessType = [System.IO.FileAccess]"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$typeAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TypeExpressionAst] }, $true)
$refType = $typeAst.TypeName.GetReflectionType()

if ($refType -eq $null) {
    Write-Host "FAIL: enum reflection type was null"
    exit 1
}

if (-not $refType.IsEnum) {
    Write-Host "FAIL: resolved type was not an Enum"
    exit 1
}

if ($refType -ne [System.IO.FileAccess]) {
    Write-Host "FAIL: enum type mismatch: '$($refType.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
