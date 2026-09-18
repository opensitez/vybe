# vybe-test: powershell/ast_loop_statements/ast_loop_dowhile_label_identifier
# DoWhileStatementAst.Label stores the loop label declared before the do keyword
$code = ":spinLock do { `$val = Read-Status } while (`$val -ne 'Ready')"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$doWhileAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.DoWhileStatementAst] }, $true)

if ($doWhileAst.Label -ne "spinLock") {
    Write-Host "FAIL: expected label 'spinLock', got '$($doWhileAst.Label)'"
    exit 1
}

Write-Host "PASS"
exit 0
