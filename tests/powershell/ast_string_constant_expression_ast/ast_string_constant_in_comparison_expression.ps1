# vybe-test: powershell/ast_string_constant_expression_ast/ast_string_constant_in_comparison_expression
# A string constant as right-hand operand of comparison -eq parses as StringConstantExpressionAst
$code = "`$response.Status -eq 'Success'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$binAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.BinaryExpressionAst] }, $true)

$right = $binAst.Right

if ($right -isnot [System.Management.Automation.Language.StringConstantExpressionAst]) {
    Write-Host "FAIL: right operand was not StringConstantExpressionAst"
    exit 1
}

if ($right.Value -ne "Success") {
    Write-Host "FAIL: right operand value mismatch: '$($right.Value)'"
    exit 1
}

Write-Host "PASS"
exit 0
