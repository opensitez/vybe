# vybe-test: powershell/command_exit/exit_in_loop
for ($i=0; $i -lt 1; $i++) { exit 0 }
exit 1
