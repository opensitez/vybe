# vybe-test: powershell/using_statements/using_namespace_system_text
using namespace System.Text

$sb = [StringBuilder]::new("Hello")
[void]$sb.Append(" World")
if ($sb.ToString() -ne "Hello World") {
    Write-Host "FAIL: using namespace System.Text expected StringBuilder 'Hello World', got $($sb.ToString())"
    exit 1
}
Write-Host "PASS"
exit 0
