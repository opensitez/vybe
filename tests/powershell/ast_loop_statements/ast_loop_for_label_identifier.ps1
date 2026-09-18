# vybe-test: powershell/ast_loop_statements/ast_loop_for_label_identifier
# ForStatementAst.Label captures the label name without the leading colon
$code = ":batchLoop for (`$i = 0; `$i -lt 100; `$i++) { `$i }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$forAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ForStatementAst] }, $true)

if ($forAst.Label -ne "batchLoop") {
    Write-Host "FAIL: expected label 'batchLoop', got '$($forAst.Label)'"
    exit 1
}

Write-Host "PASS"
exit 0
