<?php
// vybe-test: php/arrays/array_combine_invalid_count_fails
// origin: languages/php/tests/php/test_arrays.rs

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

$ok = true;
try {
    array_combine(['a', 'b'], [1]);
    $ok = false;
} catch (ValueError $e) {
    echo 'value_error';
}
if ($ok) { echo 'unexpected'; }

__vybe_check(ob_get_clean(), "value_errorunexpected");
