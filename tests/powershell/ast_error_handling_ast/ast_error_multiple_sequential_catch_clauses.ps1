# vybe-test: powershell/ast_error_handling_ast/ast_error_multiple_sequential_catch_clauses
# TryStatementAst.CatchClauses accurately indexes multiple sequential catch blocks
$code = @"
try {
    1
} catch [System.TimeoutException] {
    2
} catch [System.InvalidOperationException] {
    3
} catch {
    4
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$tryAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TryStatementAst] }, $true)

if ($tryAst.CatchClauses.Count -ne 3) {
    Write-Host "FAIL: expected 3 catch clauses, got $($tryAst.CatchClauses.Count)"
    exit 1
}

# Third clause should be the catch-all
if (-not $tryAst.CatchClauses[2].IsCatchAll) {
    Write-Host "FAIL: third catch clause was not catch-all"
    exit 1
}

Write-Host "PASS"
exit 0
