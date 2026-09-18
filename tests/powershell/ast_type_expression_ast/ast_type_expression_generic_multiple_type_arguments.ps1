# vybe-test: powershell/ast_type_expression_ast/ast_type_expression_generic_multiple_type_arguments
# Generic Dictionary type has 2 distinct generic arguments
$code = "`$dict = [System.Collections.Generic.Dictionary[string, int]]"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$typeAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TypeExpressionAst] }, $true)
$tn = $typeAst.TypeName

if ($tn.GenericArguments.Count -ne 2) {
    Write-Host "FAIL: expected 2 generic arguments, got $($tn.GenericArguments.Count)"
    exit 1
}

$keyArg = $tn.GenericArguments[0].FullName
$valArg = $tn.GenericArguments[1].FullName

if ($keyArg -ne "string" -or $valArg -ne "int") {
    Write-Host "FAIL: generic arguments mismatch: '$keyArg', '$valArg'"
    exit 1
}

Write-Host "PASS"
exit 0
