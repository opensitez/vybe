# vybe-test: powershell/ast_extent_source_mapping/ast_extent_string_literal_includes_quote_delimiters
# StringConstantExpressionAst.Extent includes the enclosing single or double quote delimiters
$code = "`$str = 'hello literal world'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$strAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)
$strExtentText = $strAst.Extent.Text

if ($strExtentText -ne "'hello literal world'") {
    Write-Host "FAIL: string literal extent did not include quotes, got: '$strExtentText'"
    exit 1
}

# Whereas .Value contains the stripped payload
if ($strAst.Value -ne "hello literal world") {
    Write-Host "FAIL: string constant Value mismatch: '$($strAst.Value)'"
    exit 1
}

Write-Host "PASS"
exit 0
