# vybe-test: powershell/ast_extent_source_mapping/ast_extent_text_matches_substring_from_offsets
# Substring of raw script using StartOffset and length matches the Extent.Text property exactly
$code = "Get-Process | Where-Object { `$_.CPU -gt 10 } | Select-Object -First 5"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$cmdAsts = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)

foreach ($cmd in $cmdAsts) {
    $extent = $cmd.Extent
    $sliceLen = $extent.EndOffset - $extent.StartOffset
    $rawSlice = $code.Substring($extent.StartOffset, $sliceLen)

    if ($rawSlice -ne $extent.Text) {
        Write-Host "FAIL: raw substring mismatch for '$($extent.Text)'"
        exit 1
    }
}

Write-Host "PASS"
exit 0
