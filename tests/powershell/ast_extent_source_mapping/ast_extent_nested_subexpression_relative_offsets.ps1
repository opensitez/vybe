# vybe-test: powershell/ast_extent_source_mapping/ast_extent_nested_subexpression_relative_offsets
# A subexpression inside an expandable string has extents strictly bounded within the string extent
$code = "`$greeting = `"Hello, `$(`$name.ToUpper())! Today is `$day.`""

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$expStringAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ExpandableStringExpressionAst] }, $true)
$subExprAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SubExpressionAst] }, $true)

$outerExtent = $expStringAst.Extent
$innerExtent = $subExprAst.Extent

if ($innerExtent.StartOffset -le $outerExtent.StartOffset) {
    Write-Host "FAIL: inner StartOffset not strictly greater than outer StartOffset"
    exit 1
}

if ($innerExtent.EndOffset -ge $outerExtent.EndOffset) {
    Write-Host "FAIL: inner EndOffset not strictly less than outer EndOffset"
    exit 1
}

if ($innerExtent.Text -ne "`$(`$name.ToUpper())") {
    Write-Host "FAIL: subexpression extent text mismatch: '$($innerExtent.Text)'"
    exit 1
}

Write-Host "PASS"
exit 0
