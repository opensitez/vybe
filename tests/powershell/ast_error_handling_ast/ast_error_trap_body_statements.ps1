# vybe-test: powershell/ast_error_handling_ast/ast_error_trap_body_statements
# TrapStatementAst.Body contains the executable statements inside the trap handler block
$code = @"
trap {
    Write-Error "Handled"
    continue
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$trapAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TrapStatementAst] }, $true)

if ($trapAst.Body.Statements.Count -ne 2) {
    Write-Host "FAIL: expected 2 statements in trap body, got $($trapAst.Body.Statements.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
