# vybe-test: powershell/command_exit/exit_in_switch
switch (1) { 1 { exit 0 } }
exit 1
