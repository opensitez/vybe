# vybe-test: powershell/ast_type_expression_ast/ast_type_expression_extent_encloses_brackets
# TypeExpressionAst.Extent.Text starts with '[' and terminates with ']'
$code = "`$targetType = [System.Text.Encoding]"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$typeAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TypeExpressionAst] }, $true)
$extentText = $typeAst.Extent.Text

if (-not $extentText.StartsWith("[")) {
    Write-Host "FAIL: TypeExpressionAst extent did not start with '[': '$extentText'"
    exit 1
}

if (-not $extentText.EndsWith("]")) {
    Write-Host "FAIL: TypeExpressionAst extent did not end with ']': '$extentText'"
    exit 1
}

if ($extentText -ne "[System.Text.Encoding]") {
    Write-Host "FAIL: extent text mismatch: '$extentText'"
    exit 1
}

Write-Host "PASS"
exit 0
