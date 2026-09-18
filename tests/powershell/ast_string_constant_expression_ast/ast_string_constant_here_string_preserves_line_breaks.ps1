# vybe-test: powershell/ast_string_constant_expression_ast/ast_string_constant_here_string_preserves_line_breaks
# Single-quoted here-strings preserve multiline breaks exactly in .Value
$code = "`$doc = @'`nParagraph 1`nParagraph 2`nParagraph 3`n'@"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$strAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)

$lines = $strAst.Value -split "`r?`n"

if ($lines.Count -ne 3) {
    Write-Host "FAIL: expected 3 lines in here-string value, got $($lines.Count)"
    exit 1
}

if ($lines[0] -ne "Paragraph 1" -or $lines[2] -ne "Paragraph 3") {
    Write-Host "FAIL: lines content mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
