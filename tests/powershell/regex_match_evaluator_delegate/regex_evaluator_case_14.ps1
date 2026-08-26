# vybe-test: powershell/regex_match_evaluator_delegate/regex_evaluator_case_14
$eval = [System.Text.RegularExpressions.MatchEvaluator]{ param($m) return "[${m}]" }
$res = [regex]::Replace("word_14", "word_14", $eval)
if ($res -ne "[word_14]") { Write-Host "FAIL: MatchEvaluator failed"; exit 1 }
Write-Host "PASS"; exit 0
