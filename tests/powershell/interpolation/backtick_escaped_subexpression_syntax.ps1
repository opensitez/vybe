# vybe-test: powershell/interpolation/backtick_escaped_subexpression_syntax
# Escaping the dollar sign of a subexpression with a backtick prevents evaluation of the subexpression
$item = "processor"

$output = "Escaped: `$(Get-Date) and literal: `$item"

if ($output -ne 'Escaped: $(Get-Date) and literal: $item') {
    Write-Host "FAIL: expected literal string with escaped dollars, got '$output'"
    exit 1
}

Write-Host "PASS"
exit 0
