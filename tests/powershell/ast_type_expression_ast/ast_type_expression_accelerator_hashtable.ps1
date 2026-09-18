# vybe-test: powershell/ast_type_expression_ast/ast_type_expression_accelerator_hashtable
# [hashtable] parses as TypeExpressionAst and resolves to System.Collections.Hashtable
$code = "`$hashType = [hashtable]"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$typeAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TypeExpressionAst] }, $true)
$refType = $typeAst.TypeName.GetReflectionType()

if ($refType -ne [System.Collections.Hashtable]) {
    Write-Host "FAIL: reflection type mismatch, expected Hashtable, got '$($refType.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
