<?php
// vybe-test: php/array_functions/array_filter_with_flag_both
// origin: languages/php/tests/php/test_array_functions.rs

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

$m = ['alpha' => 1, 'beta' => 0, 'gamma' => 2];
$b = array_filter($m, fn($v, $k) => $v === 0 || $k === 'gamma', ARRAY_FILTER_USE_BOTH);
echo implode(',', array_keys($b)) . '|' . implode(',', $b);

__vybe_check(ob_get_clean(), "beta,gamma|0,2");
