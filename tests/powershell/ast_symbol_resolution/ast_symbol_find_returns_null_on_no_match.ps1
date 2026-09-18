# vybe-test: powershell/ast_symbol_resolution/ast_symbol_find_returns_null_on_no_match
# Ast.Find returns $null when no nodes in the tree satisfy the predicate
$code = "`$counter = 1"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

# Search for a TryStatementAst which does not exist in the code
$tryNode = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TryStatementAst] }, $true)

if ($tryNode -ne $null) {
    Write-Host "FAIL: Find unexpectedly matched a non-existent TryStatementAst"
    exit 1
}

Write-Host "PASS"
exit 0
