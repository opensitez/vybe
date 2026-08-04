# vybe-test: powershell/variable_promotion/for_loop_promotion
for ($i=0; $i -lt 1; $i++) { $x = 1 }
if ($x -eq 1) { exit 0 }
exit 1
