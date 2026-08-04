# vybe-test: powershell/script_scoping/private_scope
function Test { $private:a = 'x' }
Test
if ($private:a -eq 'x') { exit 0 }
exit 1
