<?php
// vybe-test: php/array_functions_extra/array_map_with_key_mode_not_supported_keeps_single_column
// origin: languages/php/tests/php/test_array_functions_extra.rs

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

$a = [1, 2, 3];
$mapped = array_map(fn($value) => $value * 2, $a);
echo implode(',', $mapped);

__vybe_check(ob_get_clean(), "2,4,6");
