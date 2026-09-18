# vybe-test: powershell/ast_param_block_ast/ast_param_block_multiple_parameters_sequential_order
# Parameters in ParamBlockAst.Parameters are ordered in precise lexical sequence
$code = "param(`$FirstParam, `$SecondParam, `$ThirdParam, `$FourthParam, `$FifthParam)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$pb = $ast.ParamBlock
$names = @($pb.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })

$expected = @("FirstParam", "SecondParam", "ThirdParam", "FourthParam", "FifthParam")

if ($names.Count -ne 5) {
    Write-Host "FAIL: expected 5 parameters, got $($names.Count)"
    exit 1
}

for ($i = 0; $i -lt 5; $i++) {
    if ($names[$i] -ne $expected[$i]) {
        Write-Host "FAIL: at index ${i}, expected '$($expected[$i])', got '$($names[$i])'"
        exit 1
    }
}

Write-Host "PASS"
exit 0
