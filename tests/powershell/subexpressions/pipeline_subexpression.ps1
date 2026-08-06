# vybe-test: powershell/subexpressions/pipeline_subexpression
# `$( … )` collects the output of a PIPELINE, including one whose first segment
# is a value rather than a command.
$r = $(1..3 | ForEach-Object { $_ * 2 })
if ($r.Count -ne 3) {
    Write-Host "FAIL: expected 3 items, got [$($r.Count)]"
    exit 1
}
if ($r[2] -ne 6) {
    Write-Host "FAIL: expected 6, got [$($r[2])]"
    exit 1
}
Write-Host 'PASS'
exit 0
