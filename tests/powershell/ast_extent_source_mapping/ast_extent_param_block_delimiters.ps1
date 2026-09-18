# vybe-test: powershell/ast_extent_source_mapping/ast_extent_param_block_delimiters
# ParamBlockAst.Extent starts with 'param' and ends with closing parenthesis ')'
$code = "param([string]`$Name, [int]`$Age)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$paramBlockAst = $ast.ParamBlock
$extentText = $paramBlockAst.Extent.Text.Trim()

if (-not $extentText.StartsWith("param")) {
    Write-Host "FAIL: param block extent did not start with 'param': '$extentText'"
    exit 1
}

if (-not $extentText.EndsWith(")")) {
    Write-Host "FAIL: param block extent did not end with ')': '$extentText'"
    exit 1
}

Write-Host "PASS"
exit 0
