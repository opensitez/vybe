# vybe-test: powershell/comment_syntax_suite/multiline_comment_nested_like_text
<# outer<#inner#> outer #>
$value = 13
if ($value -ne 13) {
    Write-Host "FAIL: nested-like markers inside block comment changed semantics"
    exit 1
}

Write-Host 'PASS'
exit 0
