# vybe-test: powershell/regex_match_evaluator_delegate/regex_evaluator_case_4
$eval = [System.Text.RegularExpressions.MatchEvaluator]{ param($m) return "[${m}]" }
$res = [regex]::Replace("word_4", "word_4", $eval)
if ($res -ne "[word_4]") { Write-Host "FAIL: MatchEvaluator failed"; exit 1 }
Write-Host "PASS"; exit 0
