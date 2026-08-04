# vybe-test: powershell/script_scoping/script_block_scope
$script:a = 'outer'
& { $a = 'inner' }
if ($a -eq 'outer') { exit 0 }
exit 1
