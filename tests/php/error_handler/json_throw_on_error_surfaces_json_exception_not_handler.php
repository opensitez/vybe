<?php
// vybe-test: php/error_handler/json_throw_on_error_surfaces_json_exception_not_handler
// origin: languages/php/tests/php/test_error_handler.rs

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

$handler = false;
set_error_handler(function() use (&$handler): bool { $handler = true; return true; });
try { json_decode('{', flags: JSON_THROW_ON_ERROR); }
catch (JsonException $e) { echo 'json'; }
restore_error_handler();

__vybe_check(ob_get_clean(), "json");
