# vybe-test: powershell/exceptions_trap_statement_scope/trap_ast_validation
$ast = [System.Management.Automation.Language.Parser]::ParseInput('
function Test-AstTrap {
    trap { continue }
}
', [ref]$null, [ref]$null)
$trapAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TrapStatementAst] }, $true)
if ($trapAst -eq $null) {
    Write-Host "FAIL: TrapStatementAst AST validation failed"
    exit 1
}
Write-Host "PASS"
exit 0
