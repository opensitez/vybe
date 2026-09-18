# vybe-test: powershell/ast_switch_statement_ast/ast_switch_default_statements_count
# Statements declared inside default {} accurately populate SwitchStatementAst.Default.Statements
$code = @"
switch (`$val) {
    'valid' { 1 }
    default {
        `$unrecognized = `$val
        Log-Unknown `$unrecognized
        return `$false
    }
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)

if ($switchAst.Default.Statements.Count -ne 3) {
    Write-Host "FAIL: expected 3 statements in default block, got $($switchAst.Default.Statements.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
