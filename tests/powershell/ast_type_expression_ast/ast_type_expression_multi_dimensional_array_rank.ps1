# vybe-test: powershell/ast_type_expression_ast/ast_type_expression_multi_dimensional_array_rank
# Multi-dimensional array type [int[,,]] has IsArray = true and Rank = 3
$code = "`$cube = [int[,,]]"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$typeAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TypeExpressionAst] }, $true)
$tn = $typeAst.TypeName

if (-not $tn.IsArray) {
    Write-Host "FAIL: IsArray was false for multi-dimensional array"
    exit 1
}

if ($tn.Rank -ne 3) {
    Write-Host "FAIL: expected Rank 3, got $($tn.Rank)"
    exit 1
}

if ($tn.ElementType.FullName -ne "int") {
    Write-Host "FAIL: ElementType FullName mismatch, expected 'int', got '$($tn.ElementType.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
