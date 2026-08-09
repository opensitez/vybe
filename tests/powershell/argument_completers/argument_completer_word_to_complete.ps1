# vybe-test: powershell/argument_completers/argument_completer_word_to_complete
$completer = {
    param($cmd, $param, $word, $ast, $bound)
    @("Alpha", "Alpine", "Beta") | Where-Object { $_ -like "$word*" }
}
$res = @(&$completer "Cmd" "Param" "Al" $null $null)
if ($res.Count -ne 2 -or $res[0] -ne "Alpha" -or $res[1] -ne "Alpine") {
    Write-Host "FAIL: wordToComplete filter expected Alpha, Alpine, got $($res -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
