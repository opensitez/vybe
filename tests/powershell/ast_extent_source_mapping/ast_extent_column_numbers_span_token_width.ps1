# vybe-test: powershell/ast_extent_source_mapping/ast_extent_column_numbers_span_token_width
# A token extent's column width (EndColumnNumber - StartColumnNumber) equals the character length of the token
$code = "    `$myCustomVariableName = 42"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$varAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)
$extent = $varAst.Extent

$colSpan = $extent.EndColumnNumber - $extent.StartColumnNumber
$textLen = $extent.Text.Length

if ($colSpan -ne $textLen) {
    Write-Host "FAIL: column span ($colSpan) did not equal text length ($textLen)"
    exit 1
}

if ($extent.StartColumnNumber -ne 5) {
    Write-Host "FAIL: expected StartColumnNumber 5, got $($extent.StartColumnNumber)"
    exit 1
}

Write-Host "PASS"
exit 0
