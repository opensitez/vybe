<?php
// vybe-test: php/match_edge_cases/match_multiple_arms_with_precedence_guard
// origin: languages/php/tests/php/test_match_edge_cases.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

$score = 72;
echo match (true) {
    $score >= 90 => 'A',
    ($score >= 80 && $score < 90) => 'B',
    ($score >= 70 && $score < 80) => 'C',
    $score >= 60 => 'D',
    default => 'F',
};

__vybe_check(ob_get_clean(), "C");
