<?php
// vybe-test: php/array_functions_extra/array_unique_preserves_keys
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

$a = array_unique([3=>1, 5=>2, 7=>1]);
echo implode(',', array_keys($a));

__vybe_check(ob_get_clean(), "3,5");
