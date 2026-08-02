<?php
// vybe-test: php/php_expressions_ternary_coalescing_match/test_php_match_subject_with_logical_precedence_and_falsey
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

$value = 0;
echo match (true) {
    ($value > 0 && $value < 5) => 'small-positive',
    ($value === 0 && $value >= 0) => 'zero',
    default => 'other',
};

__vybe_check(ob_get_clean(), "zero");
