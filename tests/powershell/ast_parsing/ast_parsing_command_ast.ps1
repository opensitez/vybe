# vybe-test: powershell/ast_parsing/ast_parsing_command_ast
$sb = { Write-Host "Hello" }
$cmd = $sb.Ast.Find({ param($ast) $ast -is [System.Management.Automation.Language.CommandAst] }, $true)
if ($cmd.GetCommandName() -ne "Write-Host") {
    Write-Host "FAIL: CommandAst GetCommandName expected 'Write-Host', got '$($cmd.GetCommandName())'"
    exit 1
}
Write-Host "PASS"
exit 0
