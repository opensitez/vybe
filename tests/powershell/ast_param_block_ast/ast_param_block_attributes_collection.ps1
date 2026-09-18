# vybe-test: powershell/ast_param_block_ast/ast_param_block_attributes_collection
# ParamBlockAst.Attributes captures block-level attributes like CmdletBinding and OutputType
$code = @"
[CmdletBinding()]
[OutputType([string])]
param(`$Item)
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$pb = $ast.ParamBlock

if ($pb.Attributes.Count -ne 2) {
    Write-Host "FAIL: expected 2 param block attributes, got $($pb.Attributes.Count)"
    exit 1
}

$attrNames = @($pb.Attributes | ForEach-Object { $_.TypeName.FullName })
if (-not $attrNames.Contains("CmdletBinding") -or -not $attrNames.Contains("OutputType")) {
    Write-Host "FAIL: attributes mismatch: $($attrNames -join ', ')"
    exit 1
}

Write-Host "PASS"
exit 0
