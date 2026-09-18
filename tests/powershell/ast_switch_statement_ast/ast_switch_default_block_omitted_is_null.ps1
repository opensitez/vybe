# vybe-test: powershell/ast_switch_statement_ast/ast_switch_default_block_omitted_is_null
# Omitting the default block leaves SwitchStatementAst.Default as $null
$code = "switch (`$option) { 'A' { 1 } 'B' { 2 } }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)

if ($switchAst.Default -ne $null) {
    Write-Host "FAIL: Default was not null when omitted"
    exit 1
}

Write-Host "PASS"
exit 0
