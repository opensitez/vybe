# vybe-test: powershell/ast_error_handling_ast/ast_error_try_statement_extent_encloses_all_clauses
# TryStatementAst.Extent starts with 'try' and encloses try, catch, and finally blocks
$code = @"
try {
    Write-Host "work"
} catch {
    Write-Host "err"
} finally {
    Write-Host "done"
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$tryAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TryStatementAst] }, $true)
$text = $tryAst.Extent.Text.Trim()

if (-not $text.StartsWith("try")) {
    Write-Host "FAIL: try extent did not start with 'try'"
    exit 1
}

if (-not $text.Contains("catch") -or -not $text.Contains("finally")) {
    Write-Host "FAIL: try extent did not contain catch and finally keywords"
    exit 1
}

if (-not $text.EndsWith("}")) {
    Write-Host "FAIL: try extent did not end with closing brace"
    exit 1
}

Write-Host "PASS"
exit 0
