# vybe-test: powershell/implicit_returns/from_scriptblock
$block = { 'value' }
if ((& $block) -eq 'value') { exit 0 }
exit 1
