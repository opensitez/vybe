# vybe-test: powershell/ast_loop_statements/ast_loop_for_body_statements_count
# Statements inside the for body block populate ForStatementAst.Body.Statements
$code = @"
for (`$i = 0; `$i -lt 5; `$i++) {
    `$x = `$i * 2
    `$y = `$x + 1
    Write-Output `$y
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$forAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ForStatementAst] }, $true)

if ($forAst.Body.Statements.Count -ne 3) {
    Write-Host "FAIL: expected 3 statements in for body, got $($forAst.Body.Statements.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
