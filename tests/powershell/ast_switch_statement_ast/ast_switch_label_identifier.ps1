# vybe-test: powershell/ast_switch_statement_ast/ast_switch_label_identifier
# SwitchStatementAst.Label captures the label name without the leading colon
$code = ":routePacket switch (`$header) { 'TCP' { 1 } }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)

if ($switchAst.Label -ne "routePacket") {
    Write-Host "FAIL: expected label 'routePacket', got '$($switchAst.Label)'"
    exit 1
}

Write-Host "PASS"
exit 0
