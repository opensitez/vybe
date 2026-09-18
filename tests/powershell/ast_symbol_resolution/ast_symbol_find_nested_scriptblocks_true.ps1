# vybe-test: powershell/ast_symbol_resolution/ast_symbol_find_nested_scriptblocks_true
# Ast.FindAll with searchNestedScriptBlocks = $true inspects statements inside nested script blocks
$code = @"
`$outer = 1
& {
    `$inner = 2
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$vars = @($ast.FindAll({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $true))

# Must find both $outer and $inner
if ($vars.Count -ne 2) {
    Write-Host "FAIL: expected 2 variables when searchNestedScriptBlocks is true, got $($vars.Count)"
    exit 1
}

$names = $vars | ForEach-Object { $_.VariablePath.UserPath }
if (-not $names.Contains("inner") -or -not $names.Contains("outer")) {
    Write-Host "FAIL: missing inner or outer variable in nested search"
    exit 1
}

Write-Host "PASS"
exit 0
