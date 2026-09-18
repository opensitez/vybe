# vybe-test: powershell/ast_extent_source_mapping/ast_extent_hash_literal_curly_braces_boundaries
# HashtableAst.Extent begins with '@{' and concludes with closing brace '}'
$code = @"
`$options = @{
    Host = 'localhost'
    Port = 9090
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$hashAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.HashtableAst] }, $true)
$extentText = $hashAst.Extent.Text.Trim()

if (-not $extentText.StartsWith("@{")) {
    Write-Host "FAIL: hashtable extent did not start with '@{': '$extentText'"
    exit 1
}

if (-not $extentText.EndsWith("}")) {
    Write-Host "FAIL: hashtable extent did not end with '}': '$extentText'"
    exit 1
}

Write-Host "PASS"
exit 0
