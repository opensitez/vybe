# vybe-test: powershell/ast_symbol_resolution/ast_symbol_find_nested_scriptblocks_false
# Ast.FindAll with searchNestedScriptBlocks = $false skips nodes inside nested script blocks
$code = @"
`$outer = 1
& {
    `$inner = 2
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$vars = @($ast.FindAll({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $false))

# Must find only $outer, skipping $inner
if ($vars.Count -ne 1) {
    Write-Host "FAIL: expected only 1 outer variable when searchNestedScriptBlocks is false, got $($vars.Count)"
    exit 1
}

if ($vars[0].VariablePath.UserPath -ne "outer") {
    Write-Host "FAIL: matched variable was not 'outer', got '$($vars[0].VariablePath.UserPath)'"
    exit 1
}

Write-Host "PASS"
exit 0
