# vybe-test: powershell/scriptblock_ast_visitors/visitor_visit_scriptblock
$sb = { $x = 100 }
$nodes = $sb.Ast.FindAll({ param($ast) $true }, $true)
if ($nodes.Count -gt 0) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
