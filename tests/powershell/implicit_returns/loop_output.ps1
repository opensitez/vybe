# vybe-test: powershell/implicit_returns/loop_output
function Count { 1..3 | ForEach-Object { $_ } }
if ((Count | Measure-Object).Count -eq 3) { exit 0 }
exit 1
