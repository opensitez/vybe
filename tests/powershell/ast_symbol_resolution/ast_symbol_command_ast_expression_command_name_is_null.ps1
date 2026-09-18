# vybe-test: powershell/ast_symbol_resolution/ast_symbol_command_ast_expression_command_name_is_null
# Invoking a command dynamically via an expression (& $cmd) returns null for GetCommandName
$code = "& `$dynamicCommand -Argument 'val'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$cmdAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)

$name = $cmdAst.GetCommandName()

if ($name -ne $null) {
    Write-Host "FAIL: GetCommandName was not null for dynamic command expression, got '$name'"
    exit 1
}

Write-Host "PASS"
exit 0
