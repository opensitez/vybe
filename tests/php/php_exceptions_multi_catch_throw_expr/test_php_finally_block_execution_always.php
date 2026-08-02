<?php
// vybe-test: php/php_exceptions_multi_catch_throw_expr/test_php_finally_block_execution_always
// origin: languages/php/tests/php/test_php_exceptions_multi_catch_throw_expr.rs

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
    $log[] = "try";
    throw new Exception("fail");
} catch (Exception $e) {
    $log[] = "catch";
} finally {
    $log[] = "finally";
}
echo implode("-", $log);

__vybe_check(ob_get_clean(), "try-catch-finally");
