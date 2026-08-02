<?php
// vybe-test: php/preg_functions/preg_grep_filters_array
// origin: languages/php/tests/php/test_preg_functions.rs

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

$out = preg_grep('/^[aeiou]/i', ['apple', 'dog', 'egg']);
echo implode(',', array_values($out));

__vybe_check(ob_get_clean(), "apple,egg");
