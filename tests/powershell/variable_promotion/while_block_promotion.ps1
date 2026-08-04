# vybe-test: powershell/variable_promotion/while_block_promotion
while ($false) { $x = 3 }
if ($x -eq $null) { exit 0 }
exit 1
