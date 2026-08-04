# vybe-test: powershell/variable_promotion/script_block
& { $x = 7 }
if ($x -eq 7) { exit 0 }
exit 1
