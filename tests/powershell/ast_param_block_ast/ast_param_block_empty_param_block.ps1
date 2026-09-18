# vybe-test: powershell/ast_param_block_ast/ast_param_block_empty_param_block
# An empty param() block yields a ParamBlockAst with 0 parameters
$code = "param()"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$pb = $ast.ParamBlock

if ($pb -eq $null) {
    Write-Host "FAIL: ParamBlock was null for empty param() declaration"
    exit 1
}

if ($pb.Parameters.Count -ne 0) {
    Write-Host "FAIL: expected 0 parameters, got $($pb.Parameters.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
