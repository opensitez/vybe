# vybe-test: powershell/join_string_cmdlet/join_string_format_string_template
# -FormatString applies a composite formatting template to each element before joining
$res = 1..3 | Join-String -Separator '; ' -FormatString 'Val={0}'

$expected = "Val=1; Val=2; Val=3"

if ($res -ne $expected) {
    Write-Host "FAIL: expected '$expected', got '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
