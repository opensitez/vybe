# vybe-test: powershell/ast_extent_source_mapping/ast_extent_script_position_line_and_column_objects
# StartScriptPosition and EndScriptPosition provide IScriptPosition objects with Line and ColumnNumber
$code = @"
Line 1 content
`$target = 'find me'
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$varAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)
$pos = $varAst.Extent.StartScriptPosition

if ($pos.LineNumber -ne 2) {
    Write-Host "FAIL: expected LineNumber 2, got $($pos.LineNumber)"
    exit 1
}

if ($pos.ColumnNumber -ne 1) {
    Write-Host "FAIL: expected ColumnNumber 1, got $($pos.ColumnNumber)"
    exit 1
}

if ($pos.Line.Trim() -ne "`$target = 'find me'") {
    Write-Host "FAIL: script line text mismatch: '$($pos.Line)'"
    exit 1
}

Write-Host "PASS"
exit 0
