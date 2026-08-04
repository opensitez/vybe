# vybe-test: powershell/command_exit/exit_after_return
function Test { return 1; exit 4 }
if ((Test) -eq 1) { exit 0 }
exit 1
