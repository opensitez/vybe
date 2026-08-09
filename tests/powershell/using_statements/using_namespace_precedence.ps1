# vybe-test: powershell/using_statements/using_namespace_precedence
using namespace System.Text

$utf8 = [Encoding]::UTF8
if ($utf8.HeaderName -ne "utf-8") {
    Write-Host "FAIL: [Encoding] resolution via using namespace expected utf-8"
    exit 1
}
Write-Host "PASS"
exit 0
