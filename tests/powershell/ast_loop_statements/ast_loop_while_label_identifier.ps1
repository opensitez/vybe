# vybe-test: powershell/ast_loop_statements/ast_loop_while_label_identifier
# WhileStatementAst.Label stores the loop label declared before the while keyword
$code = ":pollServer while (`$true) { break :pollServer }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$whileAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.WhileStatementAst] }, $true)

if ($whileAst.Label -ne "pollServer") {
    Write-Host "FAIL: expected label 'pollServer', got '$($whileAst.Label)'"
    exit 1
}

Write-Host "PASS"
exit 0
