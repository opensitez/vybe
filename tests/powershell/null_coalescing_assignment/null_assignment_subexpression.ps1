# vybe-test: powershell/null_coalescing_assignment/null_assignment_subexpression
$tag = $null
$str = "Tag: $( $tag ??= 'DefaultTag' )"
if ($str -ne "Tag: DefaultTag" -or $tag -ne "DefaultTag") {
    Write-Host "FAIL: ??= in subexpression expected 'Tag: DefaultTag', got '$str'"
    exit 1
}
Write-Host "PASS"
exit 0
