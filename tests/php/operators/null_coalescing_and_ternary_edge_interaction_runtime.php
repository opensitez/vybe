<?php
// vybe-test: php/operators/null_coalescing_and_ternary_edge_interaction_runtime
// origin: languages/php/tests/php/test_operators.rs

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

$value = null;
echo ($value ?? 'fallback') . '|';
$value = 0;
echo ($value ?? 'fallback') . '|';
echo ($value ?: 'fallback') . '|';
$value = '';
echo ($value ?? 'fallback') . '|';
echo ($value ?: 'fallback') . '|';
echo ((null ?? 'fallback') ?: 'end') . '|';
$value = false;
echo (($value ?? true) ?: true);

__vybe_check(ob_get_clean(), "fallback|0|fallback||fallback|fallback|1");
