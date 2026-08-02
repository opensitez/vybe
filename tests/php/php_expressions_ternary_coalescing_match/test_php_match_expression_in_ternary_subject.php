<?php
// vybe-test: php/php_expressions_ternary_coalescing_match/test_php_match_expression_in_ternary_subject
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

$grade = 88;
$status = match (true) {
    $grade >= 90 => 'top',
    $grade >= 80 => 'pass',
    default => 'fail',
};
$label = $status === 'pass' ? 'ok' : 'bad';
echo $label;

__vybe_check(ob_get_clean(), "ok");
