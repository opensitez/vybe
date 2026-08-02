<?php
// vybe-test: php/php_json_serialization_contracts/test_php_json_throw_on_error_exception
// origin: languages/php/tests/php/test_php_json_serialization_contracts.rs

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
    json_decode("{invalid json}", flags: JSON_THROW_ON_ERROR);
} catch (JsonException $e) {
    echo "JsonException: " . $e->getMessage();
}

__vybe_check(ob_get_clean(), "JsonException: Syntax error");
