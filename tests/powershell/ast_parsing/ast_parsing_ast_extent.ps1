# vybe-test: powershell/ast_parsing/ast_parsing_ast_extent
$sb = { $val = "SampleText" }
$extent = $sb.Ast.Extent
if ($extent.Text -notcontains "SampleText" -and $extent.Text -notlike "*SampleText*") {
    Write-Host "FAIL: AST Extent text expected to contain 'SampleText', got '$($extent.Text)'"
    exit 1
}
Write-Host "PASS"
exit 0
