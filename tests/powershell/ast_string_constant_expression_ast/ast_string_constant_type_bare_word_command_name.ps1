# vybe-test: powershell/ast_string_constant_expression_ast/ast_string_constant_type_bare_word_command_name
# Command names in command invocations parse as StringConstantType.BareWord
$code = "Get-Process"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$strAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)

if ($strAst.StringConstantType -ne [System.Management.Automation.Language.StringConstantType]::BareWord) {
    Write-Host "FAIL: expected BareWord for command name, got '$($strAst.StringConstantType)'"
    exit 1
}

if ($strAst.Value -ne "Get-Process") {
    Write-Host "FAIL: command name value mismatch: '$($strAst.Value)'"
    exit 1
}

Write-Host "PASS"
exit 0
