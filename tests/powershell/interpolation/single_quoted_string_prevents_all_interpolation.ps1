# vybe-test: powershell/interpolation/single_quoted_string_prevents_all_interpolation
# Single-quoted string literals treat dollar signs and subexpression syntax as literal text
$userName = "alice"

$literalText = 'User: $userName and math: $(2 + 2)'

if ($literalText -ne 'User: $userName and math: $(2 + 2)') {
    Write-Host "FAIL: unexpected interpolation in single-quoted string: '$literalText'"
    exit 1
}

Write-Host "PASS"
exit 0
