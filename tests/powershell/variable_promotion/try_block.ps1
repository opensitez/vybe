# vybe-test: powershell/variable_promotion/try_block
try { $x = 8 } catch { }
if ($x -eq 8) { exit 0 }
exit 1
