# vybe-test: powershell/ast_extent_source_mapping/ast_extent_if_statement_spans_clauses_and_else
# IfStatementAst.Extent starts with 'if' and encompasses all clauses including the terminating else block
$code = @"
if (`$true) {
    `$x = 1
} else {
    `$x = 2
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$ifAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.IfStatementAst] }, $true)
$ifText = $ifAst.Extent.Text.Trim()

if (-not $ifText.StartsWith("if")) {
    Write-Host "FAIL: if statement extent did not start with 'if'"
    exit 1
}

if (-not $ifText.EndsWith("}")) {
    Write-Host "FAIL: if statement extent did not end with '}'"
    exit 1
}

if (-not $ifText.Contains("else")) {
    Write-Host "FAIL: if statement extent did not contain 'else' block"
    exit 1
}

Write-Host "PASS"
exit 0
