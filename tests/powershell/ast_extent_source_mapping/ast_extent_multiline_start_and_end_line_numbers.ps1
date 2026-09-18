# vybe-test: powershell/ast_extent_source_mapping/ast_extent_multiline_start_and_end_line_numbers
# A multi-line block statement correctly maps StartLineNumber and EndLineNumber
$code = @"
function MultilineSample {
    `$val = 10
    return `$val
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$funcAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
$extent = $funcAst.Extent

if ($extent.StartLineNumber -ne 1) {
    Write-Host "FAIL: expected StartLineNumber 1, got $($extent.StartLineNumber)"
    exit 1
}

if ($extent.EndLineNumber -ne 4) {
    Write-Host "FAIL: expected EndLineNumber 4, got $($extent.EndLineNumber)"
    exit 1
}

Write-Host "PASS"
exit 0
