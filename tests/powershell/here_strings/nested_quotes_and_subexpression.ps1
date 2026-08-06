# vybe-test: powershell/here_strings/nested_quotes_and_subexpression
# A double-quoted here-string ends only at a line-leading `"@`, so a bare `"`
# is ordinary text — and a `$( … )` inside it may still spell its own strings.
$s = 'abc'
$here = @"
u=$("q".ToUpper())
m=$($s.Replace('a','X'))
q="quoted" stays
"@
if ($here -notmatch 'u=Q') {
    Write-Host "FAIL: nested quotes in subexpression, got [$here]"
    exit 1
}
if ($here -notmatch 'm=Xbc') {
    Write-Host "FAIL: method call with quoted arg, got [$here]"
    exit 1
}
if ($here -notmatch 'q="quoted" stays') {
    Write-Host "FAIL: bare quotes should survive, got [$here]"
    exit 1
}
Write-Host 'PASS'
exit 0
