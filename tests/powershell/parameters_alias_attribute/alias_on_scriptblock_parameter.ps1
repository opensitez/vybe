# vybe-test: powershell/parameters_alias_attribute/alias_on_scriptblock_parameter
$sb = {
    param([Alias("Txt")][string]$Text)
    return "Text:$Text"
}
$res = & $sb -Txt "hello"
if ($res -ne "Text:hello") {
    Write-Host "FAIL: ScriptBlock alias parameter failed"
    exit 1
}
Write-Host "PASS"
exit 0
