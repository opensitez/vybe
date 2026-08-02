<?php
// vybe-test: php/exception_chaining/finally_runs_even_when_rethrown
// origin: languages/php/tests/php/test_exception_chaining.rs

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

$log = [];
try {
    try {
        throw new Exception('e');
    } finally {
        $log[] = 'inner_finally';
    }
} catch (Exception $e) {
    $log[] = 'caught';
}
echo implode(',', $log);

__vybe_check(ob_get_clean(), "inner_finally,caught");
