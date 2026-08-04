# vybe-test: powershell/variable_promotion/foreach_block_promotion
foreach ($item in 1) { $x = 4 }
if ($x -eq 4) { exit 0 }
exit 1
