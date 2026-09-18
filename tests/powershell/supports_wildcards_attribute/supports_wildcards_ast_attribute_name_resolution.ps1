# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_ast_attribute_name_resolution
# The AST parser successfully identifies the SupportsWildcards attribute on a parameter AST node
$code = '
function QueryItems {
    param([SupportsWildcards()][string]$ItemPattern)
}
'

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$paramAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ParameterAst] }, $true)
$attrName = $paramAst.Attributes[0].TypeName.FullName

if ($attrName -ne "SupportsWildcards") {
    Write-Host "FAIL: expected AST attribute name 'SupportsWildcards', got '$attrName'"
    exit 1
}

Write-Host "PASS"
exit 0
