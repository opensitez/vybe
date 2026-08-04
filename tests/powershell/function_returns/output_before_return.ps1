# vybe-test: powershell/function_returns/output_before_return
function Show { 'first'; return 'second' }
if ((Show | Select-Object -First 1) -eq 'first') { exit 0 }
exit 1
