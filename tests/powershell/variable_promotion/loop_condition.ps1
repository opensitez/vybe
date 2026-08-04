# vybe-test: powershell/variable_promotion/loop_condition
do { $x = 10 } until ($true)
if ($x -eq 10) { exit 0 }
exit 1
