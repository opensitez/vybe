# vybe-test: powershell/regex_match_evaluator_delegate/regex_evaluator_case_16
$eval = [System.Text.RegularExpressions.MatchEvaluator]{ param($m) return "[${m}]" }
$res = [regex]::Replace("word_16", "word_16", $eval)
if ($res -ne "[word_16]") { Write-Host "FAIL: MatchEvaluator failed"; exit 1 }
Write-Host "PASS"; exit 0
