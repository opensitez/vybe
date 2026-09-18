# vybe-test: powershell/ast_param_block_ast/ast_param_block_attribute_positional_arguments
# AttributeAst.PositionalArguments stores positional arguments like [Alias('Host', 'Server')]
$code = "param([Alias('Host', 'Server')][string]`$Endpoint)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$paramAst = $ast.ParamBlock.Parameters[0]
$aliasAttr = $paramAst.Attributes | Where-Object { $_ -is [System.Management.Automation.Language.AttributeAst] -and $_.TypeName.FullName -eq "Alias" }

if ($aliasAttr.PositionalArguments.Count -ne 2) {
    Write-Host "FAIL: expected 2 positional arguments on Alias, got $($aliasAttr.PositionalArguments.Count)"
    exit 1
}

$firstAlias = $aliasAttr.PositionalArguments[0].Value
$secondAlias = $aliasAttr.PositionalArguments[1].Value

if ($firstAlias -ne "Host" -or $secondAlias -ne "Server") {
    Write-Host "FAIL: positional arguments mismatch: '$firstAlias', '$secondAlias'"
    exit 1
}

Write-Host "PASS"
exit 0
