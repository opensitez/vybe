<?php
// vybe-test: php/match_advanced/match_with_guarding_logical_conditions
// origin: languages/php/tests/php/test_match_advanced.rs
// vybe-test-mode: compile

$score = 92;
$v = match (true) {
    $score >= 90 && $score < 100 => 'high',
    $score >= 100 => 'perfect',
    default => 'low',
};
echo $v;
