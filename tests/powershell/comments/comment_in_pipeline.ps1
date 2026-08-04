# vybe-test: powershell/comments/comment_in_pipeline
Write-Output 'PASS' | # comment
Out-Null
exit 0
