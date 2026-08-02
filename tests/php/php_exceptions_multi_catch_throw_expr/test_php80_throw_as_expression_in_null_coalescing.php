<?php
// vybe-test: php/php_exceptions_multi_catch_throw_expr/test_php80_throw_as_expression_in_null_coalescing
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

function getHost(array $config) {
    return $config["host"] ?? throw new InvalidArgumentException("Missing host");
}

try {
    getHost([]);
} catch (InvalidArgumentException $e) {
    echo "EXPR_THROWN: " . $e->getMessage();
}

__vybe_check(ob_get_clean(), "EXPR_THROWN: Missing host");
