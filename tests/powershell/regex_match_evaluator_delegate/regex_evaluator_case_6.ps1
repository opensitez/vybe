# vybe-test: powershell/regex_match_evaluator_delegate/regex_evaluator_case_6
$eval = [System.Text.RegularExpressions.MatchEvaluator]{ param($m) return "[${m}]" }
$res = [regex]::Replace("word_6", "word_6", $eval)
if ($res -ne "[word_6]") { Write-Host "FAIL: MatchEvaluator failed"; exit 1 }
Write-Host "PASS"; exit 0
