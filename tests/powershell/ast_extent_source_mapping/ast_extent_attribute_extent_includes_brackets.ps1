# vybe-test: powershell/ast_extent_source_mapping/ast_extent_attribute_extent_includes_brackets
# AttributeAst.Extent includes the enclosing square brackets '[' and ']'
$code = "param([Parameter(Mandatory = `$true)][string]`$Server)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$attrAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AttributeAst] }, $true)
$attrText = $attrAst.Extent.Text

if (-not $attrText.StartsWith("[")) {
    Write-Host "FAIL: attribute extent did not start with '[': '$attrText'"
    exit 1
}

if (-not $attrText.EndsWith("]")) {
    Write-Host "FAIL: attribute extent did not end with ']': '$attrText'"
    exit 1
}

Write-Host "PASS"
exit 0
