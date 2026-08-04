# vybe-test: powershell/variable_promotion/switch_block_promotion
switch (1) { 1 { $x = 5 } }
if ($x -eq 5) { exit 0 }
exit 1
