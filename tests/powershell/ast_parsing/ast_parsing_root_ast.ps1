# vybe-test: powershell/ast_parsing/ast_parsing_root_ast
$sb = { $x = 10; Write-Output $x }
$ast = $sb.Ast
if ($ast -eq $null -or -not ($ast -is [System.Management.Automation.Language.Ast])) {
    Write-Host "FAIL: ScriptBlock Ast property expected to be Ast instance"
    exit 1
}
Write-Host "PASS"
exit 0
