# vybe-test: powershell/ast_extent_source_mapping/ast_extent_variable_expression_includes_dollar_prefix
# VariableExpressionAst.Extent includes the leading dollar sign '$'
$code = "`$sampleTargetVariable = 'active'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$varAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)
$varText = $varAst.Extent.Text

if (-not $varText.StartsWith("`$")) {
    Write-Host "FAIL: variable extent did not start with '$': '$varText'"
    exit 1
}

if ($varText -ne "`$sampleTargetVariable") {
    Write-Host "FAIL: variable extent text mismatch: '$varText'"
    exit 1
}

Write-Host "PASS"
exit 0
