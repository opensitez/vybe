# vybe-test: powershell/script_invocation/call_scriptblock
$block = { param($x) $x + 1 }
if ((& $block 4) -eq 5) { exit 0 }
exit 1
