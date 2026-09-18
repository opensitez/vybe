# vybe-test: powershell/ast_string_constant_expression_ast/ast_string_constant_type_bare_word_argument
# Unquoted command arguments parse as StringConstantType.BareWord
$code = "Stop-Process pwsh"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$cmdAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)
$argAst = $cmdAst.CommandElements[1]

if ($argAst -isnot [System.Management.Automation.Language.StringConstantExpressionAst]) {
    Write-Host "FAIL: argument was not StringConstantExpressionAst"
    exit 1
}

if ($argAst.StringConstantType -ne [System.Management.Automation.Language.StringConstantType]::BareWord) {
    Write-Host "FAIL: expected BareWord argument, got '$($argAst.StringConstantType)'"
    exit 1
}

if ($argAst.Value -ne "pwsh") {
    Write-Host "FAIL: argument value mismatch: '$($argAst.Value)'"
    exit 1
}

Write-Host "PASS"
exit 0
