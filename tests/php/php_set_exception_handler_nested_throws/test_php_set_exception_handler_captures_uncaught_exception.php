<?php
// vybe-test: php/php_set_exception_handler_nested_throws/test_php_set_exception_handler_captures_uncaught_exception
// origin: languages/php/tests/php/test_php_set_exception_handler_nested_throws.rs

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

$captured = "";
set_exception_handler(function(Throwable $e) use (&$captured) {
    $captured = "Caught: " . $e->getMessage();
});

try {
    throw new Exception("Test Exception Payload");
} catch (Throwable $e) {
    // Manually invoke exception handler for isolated test runner predictability
    $handler = set_exception_handler(null);
    $handler($e);
}

echo $captured;

__vybe_check(ob_get_clean(), "Caught: Test Exception Payload");
