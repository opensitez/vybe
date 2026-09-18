# vybe-test: powershell/ast_extent_source_mapping/ast_extent_binary_expression_spans_operands_and_operator
# BinaryExpressionAst.Extent spans from the start of Left operand to the end of Right operand
$code = "`$result = (`$alpha + `$beta)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$binAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.BinaryExpressionAst] }, $true)
$extentText = $binAst.Extent.Text

if ($extentText -ne "`$alpha + `$beta") {
    Write-Host "FAIL: binary expression extent mismatch, expected '`$alpha + `$beta', got '$extentText'"
    exit 1
}

if ($binAst.Left.Extent.Text -ne "`$alpha" -or $binAst.Right.Extent.Text -ne "`$beta") {
    Write-Host "FAIL: child operand extents mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
