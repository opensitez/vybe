# vybe-test: powershell/implicit_returns/pipeline_result
if ((1..3 | ForEach-Object { $_ }) | Measure-Object).Count -eq 3 { exit 0 }
exit 1
