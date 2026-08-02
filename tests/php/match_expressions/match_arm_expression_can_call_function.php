<?php
// vybe-test: php/match_expressions/match_arm_expression_can_call_function
// origin: languages/php/tests/php/test_match_expressions.rs

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

function label(int $n): string { return "n=$n"; }
echo match (3) { 1 => label(1), 3 => label(3), default => 'z' };

__vybe_check(ob_get_clean(), "n=3");
