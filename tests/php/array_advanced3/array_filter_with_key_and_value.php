<?php
// vybe-test: php/array_advanced3/array_filter_with_key_and_value
// origin: languages/php/tests/php/test_array_advanced3.rs

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

$a = ['x' => 1, 'y' => 2, 'z' => 3];
$b = array_filter($a, fn($v, $k) => $v > 1 && $k !== 'z', ARRAY_FILTER_USE_BOTH);
echo implode(',', array_keys($b));

__vybe_check(ob_get_clean(), "y");
