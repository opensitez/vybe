# vybe-test: powershell/ast_type_expression_ast/ast_type_expression_reflection_type_resolution
# TypeName.GetReflectionType() successfully resolves standard types like [string] to System.String
$code = "`$t = [string]"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$typeAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TypeExpressionAst] }, $true)
$resolved = $typeAst.TypeName.GetReflectionType()

if ($resolved -ne [System.String]) {
    Write-Host "FAIL: expected System.String, got '$($resolved.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
