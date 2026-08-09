# vybe-test: powershell/argument_completers/argument_completer_command_ast
$astSample = { Test-Cmd -Arg 1 }.Ast
$completer = {
    param($cmd, $param, $word, $ast, $bound)
    if ($ast -ne $null) { "AST_PRESENT" } else { "AST_NULL" }
}
$res = &$completer "Test-Cmd" "Arg" "" $astSample $null
if ($res -ne "AST_PRESENT") {
    Write-Host "FAIL: commandAst argument handling expected AST_PRESENT"
    exit 1
}
Write-Host "PASS"
exit 0
