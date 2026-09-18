# vybe-test: powershell/ast_symbol_resolution/ast_symbol_find_first_matching_node
# Ast.Find returns the first AST node satisfying the predicate delegate
$code = @"
`$first = 10
`$second = 20
`$third = 30
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$node = $ast.Find({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)

if ($node -eq $null) {
    Write-Host "FAIL: Find returned null"
    exit 1
}

if ($node.VariablePath.UserPath -ne "first") {
    Write-Host "FAIL: expected first node 'first', got '$($node.VariablePath.UserPath)'"
    exit 1
}

Write-Host "PASS"
exit 0
