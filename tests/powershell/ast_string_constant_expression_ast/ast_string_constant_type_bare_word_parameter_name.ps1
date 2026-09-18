# vybe-test: powershell/ast_string_constant_expression_ast/ast_string_constant_type_bare_word_parameter_name
# Parameter names in command invocations parse as StringConstantType.BareWord
$code = "Get-Service -Name 'wuauserv'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$cmdAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)
$paramAst = $cmdAst.CommandElements | Where-Object { $_ -is [System.Management.Automation.Language.CommandParameterAst] }

$paramNameAst = $paramAst.Extent.Text

if (-not $paramNameAst.StartsWith("-Name")) {
    Write-Host "FAIL: parameter text mismatch: '$paramNameAst'"
    exit 1
}

Write-Host "PASS"
exit 0
