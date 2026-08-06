# vybe-test: powershell/variables/braced_name_with_space
# `${…}` delimits the NAME, so a variable name may contain characters that
# would otherwise end it — a space among them.
${my var} = 'spaced'
if (${my var} -ne 'spaced') {
    Write-Host "FAIL: got [${my var}]"
    exit 1
}
Write-Host 'PASS'
exit 0
