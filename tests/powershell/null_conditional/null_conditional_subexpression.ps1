# vybe-test: powershell/null_conditional/null_conditional_subexpression
$person = [pscustomobject]@{ Name = "Alice" }
$str = "Hello $( ${person}?.Name )"
if ($str -ne "Hello Alice") {
    Write-Host "FAIL: null-conditional in subexpression expected 'Hello Alice', got '$str'"
    exit 1
}
Write-Host "PASS"
exit 0
