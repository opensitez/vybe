<?php
// vybe-test: php/exceptions/finally_always_runs_on_exception_runtime
// origin: languages/php/tests/php/test_exceptions.rs

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

try {
    try {
        throw new Exception('inner');
    } finally {
        echo 'inner-finally|';
    }
} catch (Exception $e) {
    echo $e->getMessage() . '|';
}
echo 'done';

__vybe_check(ob_get_clean(), "inner-finally|inner|done");
