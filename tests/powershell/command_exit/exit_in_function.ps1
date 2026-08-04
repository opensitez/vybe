# vybe-test: powershell/command_exit/exit_in_function
function Test { exit 2 }
Test
exit 1
