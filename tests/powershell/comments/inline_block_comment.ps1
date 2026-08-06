# vybe-test: powershell/comments/inline_block_comment
# A block comment can sit INSIDE an expression, between operands.
$a = 1 <# middle #> + 2
if ($a -ne 3) {
    Write-Host "FAIL: got [$a]"
    exit 1
}
Write-Host 'PASS'
exit 0
