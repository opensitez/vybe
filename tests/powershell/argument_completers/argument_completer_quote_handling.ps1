# vybe-test: powershell/argument_completers/argument_completer_quote_handling
$completer = {
    param($w)
    if ($w.StartsWith("'") -or $w.StartsWith('"')) {
        $w.Substring(1)
    } else {
        $w
    }
}
$res = &$completer "'Quoted"
if ($res -ne "Quoted") {
    Write-Host "FAIL: quoted word completion expected Quoted, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
