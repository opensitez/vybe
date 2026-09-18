# vybe-test: powershell/supports_paging_parameters/supports_paging_ast_attribute_inspection
# The PowerShell AST parser exposes SupportsPaging in the NamedArguments collection of CmdletBindingAttribute
$code = '
function QueryPagedEntities {
    [CmdletBinding(SupportsPaging = $true)]
    param()
}
'

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$funcAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
$attrAst = $funcAst.Body.ParamBlock.Attributes[0]
$namedArg = $attrAst.NamedArguments | Where-Object { $_.ArgumentName -eq "SupportsPaging" }

if ($namedArg -eq $null) {
    Write-Host "FAIL: NamedArgument 'SupportsPaging' missing from AST"
    exit 1
}

if ($namedArg.Argument.VariablePath.UserPath -ne "true") {
    Write-Host "FAIL: AST argument expression mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
