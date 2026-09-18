# vybe-test: powershell/ast_loop_statements/ast_loop_for_initializer_condition_iterator
# ForStatementAst cleanly populates Initializer, Condition, and Iterator expressions
$code = "for (`$i = 0; `$i -lt 10; `$i++) { `$i }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$forAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ForStatementAst] }, $true)

if ($forAst.Initializer -eq $null -or $forAst.Condition -eq $null -or $forAst.Iterator -eq $null) {
    Write-Host "FAIL: one or more for loop clauses was null"
    exit 1
}

if ($forAst.Initializer.Extent.Text -ne "`$i = 0") {
    Write-Host "FAIL: Initializer mismatch: '$($forAst.Initializer.Extent.Text)'"
    exit 1
}

if ($forAst.Condition.Extent.Text -ne "`$i -lt 10") {
    Write-Host "FAIL: Condition mismatch: '$($forAst.Condition.Extent.Text)'"
    exit 1
}

if ($forAst.Iterator.Extent.Text -ne "`$i++") {
    Write-Host "FAIL: Iterator mismatch: '$($forAst.Iterator.Extent.Text)'"
    exit 1
}

Write-Host "PASS"
exit 0
