# vybe-test: powershell/string_encoding_base64/base64_encoded_command_simulation
# PowerShell encoded commands use Unicode (UTF-16LE) Base64
$script = 'Write-Output "OK"'
$bytes = [System.Text.Encoding]::Unicode.GetBytes($script)
$b64 = [System.Convert]::ToBase64String($bytes)
$decoded = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($b64))
if ($decoded -ne $script) {
    Write-Host "FAIL: UTF-16LE Base64 simulation failed"
    exit 1
}
Write-Host "PASS"
exit 0
