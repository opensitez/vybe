# vybe-test: powershell/ast_loop_statements/ast_loop_continue_statement_unlabeled
# Bare unlabeled continue statement has ContinueStatementAst.Label equal to $null
$code = "foreach (`$item in `$list) { if (`$item -eq 0) { continue } }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$contAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ContinueStatementAst] }, $true)

if ($contAst.Label -ne $null) {
    Write-Host "FAIL: unlabeled continue had non-null Label"
    exit 1
}

Write-Host "PASS"
exit 0
