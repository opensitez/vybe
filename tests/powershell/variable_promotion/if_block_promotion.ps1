# vybe-test: powershell/variable_promotion/if_block_promotion
if ($true) { $x = 2 }
if ($x -eq 2) { exit 0 }
exit 1
