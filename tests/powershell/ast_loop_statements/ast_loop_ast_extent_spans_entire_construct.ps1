# vybe-test: powershell/ast_loop_statements/ast_loop_ast_extent_spans_entire_construct
# ForEachStatementAst.Extent starts with the label or foreach keyword and concludes at the closing brace
$code = @"
:mainLoop foreach (`$x in `$items) {
    `$x * 2
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$foreachAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ForEachStatementAst] }, $true)
$extentText = $foreachAst.Extent.Text.Trim()

if (-not $extentText.StartsWith(":mainLoop")) {
    Write-Host "FAIL: extent did not start with label: '$extentText'"
    exit 1
}

if (-not $extentText.EndsWith("}")) {
    Write-Host "FAIL: extent did not end with closing brace: '$extentText'"
    exit 1
}

Write-Host "PASS"
exit 0
