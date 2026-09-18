# vybe-test: powershell/ast_param_block_ast/ast_param_block_parameters_count
# ParamBlockAst.Parameters.Count accurately reflects the number of parameters declared
$code = "param(`$Alpha, `$Beta, `$Gamma, `$Delta)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$pb = $ast.ParamBlock

if ($pb.Parameters.Count -ne 4) {
    Write-Host "FAIL: expected 4 parameters, got $($pb.Parameters.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
