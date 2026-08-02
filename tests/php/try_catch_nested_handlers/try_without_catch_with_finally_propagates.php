<?php
// vybe-test: php/try_catch_nested_handlers/try_without_catch_with_finally_propagates
// origin: languages/php/tests/php/test_try_catch_nested_handlers.rs

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
    try { throw new Exception('up'); }
    finally { $log[] = 'finally'; }
} catch (Exception $e) {
    $log[] = 'caught';
}
echo implode(',', $log);

__vybe_check(ob_get_clean(), "finally,caught");
