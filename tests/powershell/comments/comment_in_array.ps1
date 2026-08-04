# vybe-test: powershell/comments/comment_in_array
$arr = @(1, 2, 3) # array
if ($arr[0] -eq 1) {
    Write-Output 'PASS'
}
exit 0
