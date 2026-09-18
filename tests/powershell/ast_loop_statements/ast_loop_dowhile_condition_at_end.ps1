# vybe-test: powershell/ast_loop_statements/ast_loop_dowhile_condition_at_end
# DoWhileStatementAst contains Condition evaluated post-iteration
$code = "do { `$count++ } while (`$count -lt 5)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$doWhileAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.DoWhileStatementAst] }, $true)

if ($doWhileAst -eq $null) {
    Write-Host "FAIL: DoWhileStatementAst not found"
    exit 1
}

if ($doWhileAst.Condition.Extent.Text -ne "`$count -lt 5") {
    Write-Host "FAIL: condition mismatch: '$($doWhileAst.Condition.Extent.Text)'"
    exit 1
}

Write-Host "PASS"
exit 0
