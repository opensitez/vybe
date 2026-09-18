# vybe-test: powershell/ast_extent_source_mapping/ast_extent_here_string_literal_multiline_delimiters
# Multi-line here-string AST extent begins with '@'' and terminates with '''@'
$code = @"
`$here = @'
Line one
Line two
'@
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$strAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)
$extentText = $strAst.Extent.Text.Trim()

if (-not $extentText.StartsWith("@'")) {
    Write-Host "FAIL: here-string extent did not start with @', got: '$extentText'"
    exit 1
}

if (-not $extentText.EndsWith("'@")) {
    Write-Host "FAIL: here-string extent did not end with '@, got: '$extentText'"
    exit 1
}

Write-Host "PASS"
exit 0
