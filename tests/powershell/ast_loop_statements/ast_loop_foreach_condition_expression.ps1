# vybe-test: powershell/ast_loop_statements/ast_loop_foreach_condition_expression
# ForEachStatementAst.Condition contains the collection expression being enumerated
$code = "foreach (`$item in 1..100) { `$item * 2 }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$foreachAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ForEachStatementAst] }, $true)

$conditionText = $foreachAst.Condition.Extent.Text

if ($conditionText -ne "1..100") {
    Write-Host "FAIL: condition collection mismatch, expected '1..100', got '$conditionText'"
    exit 1
}

Write-Host "PASS"
exit 0
