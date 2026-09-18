# vybe-test: powershell/pscmdlet_transaction_binding/transaction_binding_ast_named_argument_inspection
# The PowerShell AST parser exposes SupportsTransactions in the NamedArguments collection of the attribute AST
$code = '
function SampleTxnAstCheck {
    [CmdletBinding(SupportsTransactions = $true)]
    param($Param1)
}
'

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

if ($errors.Count -gt 0) {
    Write-Host "FAIL: AST parsing failed with errors: $($errors -join '; ')"
    exit 1
}

$funcAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
$attrAst = $funcAst.Body.ParamBlock.Attributes[0]

$namedArg = $attrAst.NamedArguments | Where-Object { $_.ArgumentName -eq "SupportsTransactions" }

if ($namedArg -eq $null) {
    Write-Host "FAIL: NamedArgument 'SupportsTransactions' not found in AST"
    exit 1
}

Write-Host "PASS"
exit 0
