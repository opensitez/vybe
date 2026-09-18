# vybe-test: powershell/ast_loop_statements/ast_loop_break_statement_unlabeled
# Bare unlabeled break statement has BreakStatementAst.Label equal to $null
$code = "while (`$true) { break }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$breakAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.BreakStatementAst] }, $true)

if ($breakAst.Label -ne $null) {
    Write-Host "FAIL: unlabeled break statement had non-null Label"
    exit 1
}

Write-Host "PASS"
exit 0
