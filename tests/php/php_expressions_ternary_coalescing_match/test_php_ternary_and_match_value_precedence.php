<?php
// vybe-test: php/php_expressions_ternary_coalescing_match/test_php_ternary_and_match_value_precedence
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

$input = 1;
$status = $input ?: 'falsy' ? 'A' : 'B';
$status2 = ($input ?: 'falsy') ? 'A' : 'B';
echo $status . '|' . $status2;

__vybe_check(ob_get_clean(), "1|A");
