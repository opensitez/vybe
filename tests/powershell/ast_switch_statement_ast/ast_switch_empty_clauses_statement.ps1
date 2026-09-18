# vybe-test: powershell/ast_switch_statement_ast/ast_switch_empty_clauses_statement
# An empty switch statement switch ($x) {} parses with Clauses.Count = 0 and Default = null
$code = "switch (`$x) { }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)

if ($switchAst.Clauses.Count -ne 0) {
    Write-Host "FAIL: expected 0 clauses in empty switch, got $($switchAst.Clauses.Count)"
    exit 1
}

if ($switchAst.Default -ne $null) {
    Write-Host "FAIL: Default was not null in empty switch"
    exit 1
}

Write-Host "PASS"
exit 0
