# vybe-test: powershell/ast_string_constant_expression_ast/ast_string_constant_special_characters_newline_tab
# Escaped special characters `n and `t inside double-quoted constants resolve properly in .Value
$code = '$formatted = "Column1`tColumn2`nRow1`tRow2"'

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$strAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)

if (-not $strAst.Value.Contains("`n") -or -not $strAst.Value.Contains("`t")) {
    Write-Host "FAIL: special characters newline/tab not unescaped in .Value"
    exit 1
}

Write-Host "PASS"
exit 0
