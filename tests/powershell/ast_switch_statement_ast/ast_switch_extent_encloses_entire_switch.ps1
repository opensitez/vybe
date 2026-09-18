# vybe-test: powershell/ast_switch_statement_ast/ast_switch_extent_encloses_entire_switch
# SwitchStatementAst.Extent starts with the label or switch keyword and terminates at the closing brace
$code = @"
:protocolSwitch switch (`$protocol) {
    'HTTP'  { 80 }
    'HTTPS' { 443 }
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)
$text = $switchAst.Extent.Text.Trim()

if (-not $text.StartsWith(":protocolSwitch")) {
    Write-Host "FAIL: extent did not start with label: '$text'"
    exit 1
}

if (-not $text.EndsWith("}")) {
    Write-Host "FAIL: extent did not end with closing brace: '$text'"
    exit 1
}

Write-Host "PASS"
exit 0
