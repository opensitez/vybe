<?php
// vybe-test: php/php_expressions_ternary_coalescing_match/test_php_short_ternary_on_unset_index_with_coalesce
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

$data = ['x' => null];
$x = ($data['x'] ?? 'fallback') ?: 'alt';
$y = $data['y'] ?? 'fallback';
echo $x . '|' . $y;

__vybe_check(ob_get_clean(), "alt|fallback");
