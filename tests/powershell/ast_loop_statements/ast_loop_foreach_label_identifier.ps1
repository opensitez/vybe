# vybe-test: powershell/ast_loop_statements/ast_loop_foreach_label_identifier
# ForEachStatementAst.Label stores the loop label declared before the foreach keyword
$code = ":streamLoop foreach (`$n in `$numbers) { `$n }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$foreachAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ForEachStatementAst] }, $true)

if ($foreachAst.Label -ne "streamLoop") {
    Write-Host "FAIL: expected label 'streamLoop', got '$($foreachAst.Label)'"
    exit 1
}

Write-Host "PASS"
exit 0
