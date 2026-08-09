# vybe-test: powershell/argument_completers/argument_completer_in_advanced_function
function Get-Server {
    [CmdletBinding()]
    param(
        [ArgumentCompleter({ param($c,$p,$w,$a,$b) @("web01", "db01") })]
        [string]$Name
    )
}
$param = (Get-Command Get-Server).Parameters["Name"]
if ($param.Attributes.Count -lt 1) {
    Write-Host "FAIL: advanced function parameter completer attribute missing"
    exit 1
}
Write-Host "PASS"
exit 0
