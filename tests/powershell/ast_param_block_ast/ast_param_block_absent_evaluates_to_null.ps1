# vybe-test: powershell/ast_param_block_ast/ast_param_block_absent_evaluates_to_null
# A script with no param block evaluates ast.ParamBlock as $null
$code = "Write-Host 'no param block declared here'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

if ($ast.ParamBlock -ne $null) {
    Write-Host "FAIL: ParamBlock was not null when param() was omitted"
    exit 1
}

Write-Host "PASS"
exit 0
