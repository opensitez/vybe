<?php
// vybe-test: php/array_functions/array_filter_with_flag_preserve_keys_runtime
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

$a = [0 => 'keep', 1 => '', 2 => 'x'];
$b = array_filter($a, fn($v) => $v !== '');
echo implode('|', $b);
echo ':';
echo array_key_last($b);

__vybe_check(ob_get_clean(), "keep|x:2");
