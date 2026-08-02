<?php
// vybe-test: php/php_expressions_ternary_coalescing_match/test_php_ternary_nested_with_match_subject_side_effect
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

$calls = 0;
$next = function() use (&$calls) { $calls++; return 10; };
$expr = $next() > 5 ? match (10) {
    10 => 'ten',
    default => 'other',
} : 'low';
echo $expr . '|' . $calls;

__vybe_check(ob_get_clean(), "ten|1");
