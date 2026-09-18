# vybe-test: powershell/ast_string_constant_expression_ast/ast_string_constant_in_array_literal
# String constants inside an array literal parse cleanly as StringConstantExpressionAst elements
$code = "`$roles = @('Reader', 'Contributor', 'Owner')"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$arrAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ArrayLiteralAst] }, $true)

if ($arrAst.Elements.Count -ne 3) {
    Write-Host "FAIL: expected 3 elements in array literal, got $($arrAst.Elements.Count)"
    exit 1
}

foreach ($elem in $arrAst.Elements) {
    if ($elem -isnot [System.Management.Automation.Language.StringConstantExpressionAst]) {
        Write-Host "FAIL: element was not StringConstantExpressionAst: '$($elem.GetType().Name)'"
        exit 1
    }
}

Write-Host "PASS"
exit 0
