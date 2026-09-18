# vybe-test: powershell/ast_switch_statement_ast/ast_switch_clauses_count_matches_branches
# SwitchStatementAst.Clauses.Count reflects the number of non-default evaluation branches
$code = @"
switch (`$val) {
    'alpha' { 1 }
    'beta'  { 2 }
    'gamma' { 3 }
    default { 0 }
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)

if ($switchAst.Clauses.Count -ne 3) {
    Write-Host "FAIL: expected 3 clauses, got $($switchAst.Clauses.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
