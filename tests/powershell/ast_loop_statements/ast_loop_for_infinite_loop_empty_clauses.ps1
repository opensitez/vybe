# vybe-test: powershell/ast_loop_statements/ast_loop_for_infinite_loop_empty_clauses
# An infinite for loop for (;;) evaluates Initializer, Condition, and Iterator as null
$code = "for (;;) { break }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$forAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ForStatementAst] }, $true)

if ($forAst.Initializer -ne $null) {
    Write-Host "FAIL: Initializer was not null for for(;;)"
    exit 1
}

if ($forAst.Condition -ne $null) {
    Write-Host "FAIL: Condition was not null for for(;;)"
    exit 1
}

if ($forAst.Iterator -ne $null) {
    Write-Host "FAIL: Iterator was not null for for(;;)"
    exit 1
}

Write-Host "PASS"
exit 0
