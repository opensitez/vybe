# vybe-test: powershell/regex_match_evaluator_delegate/regex_evaluator_case_18
$eval = [System.Text.RegularExpressions.MatchEvaluator]{ param($m) return "[${m}]" }
$res = [regex]::Replace("word_18", "word_18", $eval)
if ($res -ne "[word_18]") { Write-Host "FAIL: MatchEvaluator failed"; exit 1 }
Write-Host "PASS"; exit 0
