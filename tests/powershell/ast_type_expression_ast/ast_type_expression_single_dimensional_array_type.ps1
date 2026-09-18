# vybe-test: powershell/ast_type_expression_ast/ast_type_expression_single_dimensional_array_type
# Single-dimensional array type [string[]] has IsArray = true, Rank = 1, and ElementType.FullName = 'string'
$code = "`$arr = [string[]]"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$typeAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TypeExpressionAst] }, $true)
$tn = $typeAst.TypeName

if (-not $tn.IsArray) {
    Write-Host "FAIL: IsArray was false for [string[]]"
    exit 1
}

if ($tn.Rank -ne 1) {
    Write-Host "FAIL: expected Rank 1, got $($tn.Rank)"
    exit 1
}

if ($tn.ElementType.FullName -ne "string") {
    Write-Host "FAIL: ElementType FullName mismatch, expected 'string', got '$($tn.ElementType.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
