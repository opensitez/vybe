# vybe-test: powershell/ast_type_expression_ast/ast_type_expression_generic_single_type_argument
# Generic type [System.Collections.Generic.List[int]] has IsGeneric = true and 1 generic argument
$code = "`$list = [System.Collections.Generic.List[int]]"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$typeAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TypeExpressionAst] }, $true)
$tn = $typeAst.TypeName

if (-not $tn.IsGeneric) {
    Write-Host "FAIL: IsGeneric was false for generic List"
    exit 1
}

if ($tn.GenericArguments.Count -ne 1) {
    Write-Host "FAIL: expected 1 generic argument, got $($tn.GenericArguments.Count)"
    exit 1
}

if ($tn.GenericArguments[0].FullName -ne "int") {
    Write-Host "FAIL: generic argument mismatch, expected 'int', got '$($tn.GenericArguments[0].FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
