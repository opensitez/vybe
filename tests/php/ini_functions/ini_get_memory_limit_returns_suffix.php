<?php
// vybe-test: php/ini_functions/ini_get_memory_limit_returns_suffix
// origin: languages/php/tests/php/test_ini_functions.rs

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

echo str_ends_with(ini_get('memory_limit'), 'M') || str_ends_with(ini_get('memory_limit'), 'G') || ini_get('memory_limit') === '-1' ? 'limit' : 'other';

__vybe_check(ob_get_clean(), "limit");
