# vybe-test: powershell/ast_extent_source_mapping/ast_extent_single_line_statement_line_equality
# A single-line statement has identical StartLineNumber and EndLineNumber
$code = @"
`$x = 100
`$y = 200
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assigns = @($ast.FindAll({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true))

if ($assigns.Count -ne 2) {
    Write-Host "FAIL: expected 2 assignments, got $($assigns.Count)"
    exit 1
}

$first = $assigns[0].Extent
if ($first.StartLineNumber -ne 1 -or $first.EndLineNumber -ne 1) {
    Write-Host "FAIL: first assignment line numbers mismatch: $($first.StartLineNumber) - $($first.EndLineNumber)"
    exit 1
}

$second = $assigns[1].Extent
if ($second.StartLineNumber -ne 2 -or $second.EndLineNumber -ne 2) {
    Write-Host "FAIL: second assignment line numbers mismatch: $($second.StartLineNumber) - $($second.EndLineNumber)"
    exit 1
}

Write-Host "PASS"
exit 0
