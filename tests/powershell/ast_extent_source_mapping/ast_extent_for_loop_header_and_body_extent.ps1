# vybe-test: powershell/ast_extent_source_mapping/ast_extent_for_loop_header_and_body_extent
# ForStatementAst.Extent starts at 'for' and terminates after the body scriptblock
$code = "for (`$i = 0; `$i -lt 10; `$i++) { `$i * 2 }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$forAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ForStatementAst] }, $true)
$forText = $forAst.Extent.Text

if (-not $forText.StartsWith("for")) {
    Write-Host "FAIL: for loop extent did not start with 'for': '$forText'"
    exit 1
}

if (-not $forText.EndsWith("}")) {
    Write-Host "FAIL: for loop extent did not end with '}': '$forText'"
    exit 1
}

if ($forAst.Extent.StartOffset -ne 0 -or $forAst.Extent.EndOffset -ne $code.Length) {
    Write-Host "FAIL: for loop extent offsets mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
