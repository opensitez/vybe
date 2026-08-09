# vybe-test: powershell/ast_parsing/ast_parsing_hashtable_ast
$sb = { @{ Key = "Val" } }
$hashAst = $sb.Ast.Find({ param($ast) $ast -is [System.Management.Automation.Language.HashtableAst] }, $true)
if ($hashAst.KeyValuePairs.Count -ne 1) {
    Write-Host "FAIL: HashtableAst KeyValuePairs count expected 1, got $($hashAst.KeyValuePairs.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
