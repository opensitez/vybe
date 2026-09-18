# vybe-test: powershell/ast_extent_source_mapping/ast_extent_try_catch_finally_compound_extent
# TryStatementAst.Extent starts at 'try' and ends at the closing brace of the finally block
$code = @"
try {
    Write-Host "work"
} catch {
    Write-Host "err"
} finally {
    Write-Host "clean"
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$tryAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TryStatementAst] }, $true)
$tryText = $tryAst.Extent.Text.Trim()

if (-not $tryText.StartsWith("try")) {
    Write-Host "FAIL: try statement extent did not start with 'try'"
    exit 1
}

if (-not $tryText.Contains("catch") -or -not $tryText.Contains("finally")) {
    Write-Host "FAIL: try statement extent did not encompass catch and finally blocks"
    exit 1
}

if (-not $tryText.EndsWith("}")) {
    Write-Host "FAIL: try statement extent did not terminate with '}'"
    exit 1
}

Write-Host "PASS"
exit 0
