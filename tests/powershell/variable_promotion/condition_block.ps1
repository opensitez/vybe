# vybe-test: powershell/variable_promotion/condition_block
if ($true) { $x = 9 }
if ($x -eq 9) { exit 0 }
exit 1
