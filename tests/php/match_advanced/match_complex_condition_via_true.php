<?php
// vybe-test: php/match_advanced/match_complex_condition_via_true
// origin: languages/php/tests/php/test_match_advanced.rs
// vybe-test-mode: compile

$score = 87;
$grade = match(true) {
    $score >= 90 => 'A',
    $score >= 80 => 'B',
    $score >= 70 => 'C',
    $score >= 60 => 'D',
    default      => 'F',
};
echo $grade;
