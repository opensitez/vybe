# vybe-test: powershell/argument_completers/argument_completer_attribute_basic
function Test-Completer {
    param(
        [ArgumentCompleter({
            param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
            return @("Option1", "Option2")
        })]
        [string]$Choice
    )
}
$param = (Get-Command Test-Completer).Parameters["Choice"]
if ($param.Attributes.Count -lt 1) {
    Write-Host "FAIL: ArgumentCompleter attribute missing from parameter"
    exit 1
}
Write-Host "PASS"
exit 0
