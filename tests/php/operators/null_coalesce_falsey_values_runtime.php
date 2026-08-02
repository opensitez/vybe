<?php
// vybe-test: php/operators/null_coalesce_falsey_values_runtime
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

echo ('' ?? 'empty') . '|';
echo (0 ?? 'zero') . '|';
$value = null;
echo (($value ?? null) ?? 'fallback') . '|';
$value = 0;
echo ($value ?? 'fallback') . '|';
$value = '';
echo ($value ?? 'fallback') . '|';

__vybe_check(ob_get_clean(), "|0|fallback|0||");
