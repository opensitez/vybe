<?php
// vybe-test: php/php_expressions_ternary_coalescing_match/test_php_chained_ternary_with_parentheses
// origin: languages/php/tests/php/test_php_expressions_ternary_coalescing_match.rs

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

$score = 85;
$grade = ($score >= 90) ? "A" : (($score >= 80) ? "B" : "C");
echo $grade;

__vybe_check(ob_get_clean(), "B");
