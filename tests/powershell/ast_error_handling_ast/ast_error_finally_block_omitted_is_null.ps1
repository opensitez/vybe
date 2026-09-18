# vybe-test: powershell/ast_error_handling_ast/ast_error_finally_block_omitted_is_null
# TryStatementAst.Finally is $null when finally block is omitted
$code = "try { 1 } catch { 2 }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$tryAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TryStatementAst] }, $true)

if ($tryAst.Finally -ne $null) {
    Write-Host "FAIL: Finally was not null when omitted"
    exit 1
}

Write-Host "PASS"
exit 0
