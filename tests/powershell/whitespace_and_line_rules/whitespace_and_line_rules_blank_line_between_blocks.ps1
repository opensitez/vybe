# vybe-test: powershell/whitespace_and_line_rules/blank_line_between_blocks
$values = @()
if ($true) {
    $values += 1
}

if ($true) {
    $values += 1
}

if (($values.Count -ne 2) -or (($values -join ',') -ne '1,1')) {
    Write-Host 'FAIL: blank line between blocks changed execution semantics'
    exit 1
}

Write-Host 'PASS'
exit 0
