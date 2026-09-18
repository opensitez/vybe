# vybe-test: powershell/ast_loop_statements/ast_loop_dountil_label_identifier
# DoUntilStatementAst.Label stores the loop label declared before the do keyword
$code = ":retryPolicy do { Invoke-ApiCall } until (`$success)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$doUntilAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.DoUntilStatementAst] }, $true)

if ($doUntilAst.Label -ne "retryPolicy") {
    Write-Host "FAIL: expected label 'retryPolicy', got '$($doUntilAst.Label)'"
    exit 1
}

Write-Host "PASS"
exit 0
