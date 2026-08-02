<?php
// vybe-test: php/arrays/array_search_with_strict_flag
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

$a = [1, '2', 3];
echo in_array('2', $a, true) ? 'strict-true' : 'strict-false';
echo '|';
echo in_array(2, $a, true) ? 'strict-true-2' : 'strict-false-2';

__vybe_check(ob_get_clean(), "strict-true|strict-false-2");
