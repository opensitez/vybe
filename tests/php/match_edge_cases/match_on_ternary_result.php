<?php
// vybe-test: php/match_edge_cases/match_on_ternary_result
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

$n = 7;
echo match($n > 5 ? 'big' : 'small') { 'big' => 'large number', 'small' => 'tiny number' };

__vybe_check(ob_get_clean(), "large number");
