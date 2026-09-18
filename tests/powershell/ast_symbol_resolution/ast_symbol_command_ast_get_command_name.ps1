# vybe-test: powershell/ast_symbol_resolution/ast_symbol_command_ast_get_command_name
# CommandAst.GetCommandName returns the command invocation target as a string
$code = "Get-ChildItem -Path '/var/log' -Recurse"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$cmdAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)

$cmdName = $cmdAst.GetCommandName()

if ($cmdName -ne "Get-ChildItem") {
    Write-Host "FAIL: GetCommandName mismatch, expected 'Get-ChildItem', got '$cmdName'"
    exit 1
}

Write-Host "PASS"
exit 0
