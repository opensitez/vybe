# vybe-test: powershell/ast_type_expression_ast/ast_type_expression_unresolved_type_name_does_not_throw
# Parsing a non-existent or unloaded type name succeeds and accurately stores the parsed text
$code = "`$customType = [CustomEnterprise.VirtualMesh.DistributedCoordinator]"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$typeAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TypeExpressionAst] }, $true)

if ($typeAst -eq $null) {
    Write-Host "FAIL: TypeExpressionAst was null for unresolved type"
    exit 1
}

$fullName = $typeAst.TypeName.FullName
$expected = "CustomEnterprise.VirtualMesh.DistributedCoordinator"

if ($fullName -ne $expected) {
    Write-Host "FAIL: unresolved type FullName mismatch: '$fullName'"
    exit 1
}

$refType = $typeAst.TypeName.GetReflectionType()
if ($refType -ne $null) {
    Write-Host "FAIL: expected GetReflectionType to be null for unloaded type, got '$($refType.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
