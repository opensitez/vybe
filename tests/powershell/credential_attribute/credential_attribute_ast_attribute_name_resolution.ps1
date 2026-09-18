# vybe-test: powershell/credential_attribute/credential_attribute_ast_attribute_name_resolution
# The PowerShell AST parser resolves the Credential attribute name in the parameter AST
$code = '
using namespace System.Management.Automation
function TestCredAst {
    param([Credential()][pscredential]$Auth)
}
'

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$paramAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ParameterAst] }, $true)
$attrName = $paramAst.Attributes[0].TypeName.FullName

if ($attrName -ne "Credential") {
    Write-Host "FAIL: expected AST attribute name 'Credential', got '$attrName'"
    exit 1
}

Write-Host "PASS"
exit 0
