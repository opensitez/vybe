# vybe-test: powershell/dynamic_typing/pipeline_output
$x = 1
$x = $x | ForEach-Object { $_ + 1 }
if ($x -eq 2) { exit 0 }
exit 1
