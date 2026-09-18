# vybe-test: powershell/ast_extent_source_mapping/ast_extent_function_definition_boundaries
# FunctionDefinitionAst.Extent begins with the 'function' keyword and terminates with the closing brace
$code = "function MyFunction { return 123 }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$funcAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
$extentText = $funcAst.Extent.Text

if (-not $extentText.StartsWith("function")) {
    Write-Host "FAIL: function extent did not start with 'function': '$extentText'"
    exit 1
}

if (-not $extentText.EndsWith("}")) {
    Write-Host "FAIL: function extent did not end with '}': '$extentText'"
    exit 1
}

Write-Host "PASS"
exit 0
