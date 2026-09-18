# vybe-test: powershell/ast_switch_statement_ast/ast_switch_condition_pipeline_expression
# SwitchStatementAst.Condition accurately captures the evaluated target expression
$code = "switch (`$statusCode + 1) { 200 { 'OK' } }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)

if ($switchAst.Condition.Extent.Text -ne "`$statusCode + 1") {
    Write-Host "FAIL: condition expression mismatch: '$($switchAst.Condition.Extent.Text)'"
    exit 1
}

Write-Host "PASS"
exit 0
