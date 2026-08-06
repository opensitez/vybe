# vybe-test: powershell/switch/pattern_matching
# Wildcard patterns in a switch require `-Wildcard`. A PLAIN switch compares
# each label with `-eq`, so 'h*' would only match the literal string "h*".
$value = 'hello'
$result = switch -Wildcard ($value) { 'h*' { 'match' } default { 'miss' } }
if ($result -ne 'match') {
    Write-Host 'FAIL'
    exit 1
}
$plain = switch ($value) { 'h*' { 'match' } default { 'miss' } }
if ($plain -ne 'miss') {
    Write-Host 'FAIL: a plain switch must not glob'
    exit 1
}
Write-Host 'PASS'
exit 0
