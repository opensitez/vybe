# vybe-test: powershell/comments/hash_inside_string
# `#` starts a comment only in CODE. Inside a string it is an ordinary
# character, so the rest of the line is not discarded.
$s = "colour #42 selected"
if ($s -ne 'colour #42 selected') {
    Write-Host "FAIL: got [$s]"
    exit 1
}
Write-Host 'PASS'
exit 0
