# vybe-test: powershell/ast_switch_statement_ast/ast_switch_default_block_presence
# A default block populates SwitchStatementAst.Default with a non-null StatementBlockAst
$code = "switch (`$env) { 'prod' { 1 } default { 0 } }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)

if ($switchAst.Default -eq $null) {
    Write-Host "FAIL: Default was null when default block was present"
    exit 1
}

if ($switchAst.Default -isnot [System.Management.Automation.Language.StatementBlockAst]) {
    Write-Host "FAIL: expected Default to be StatementBlockAst, got '$($switchAst.Default.GetType().Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
