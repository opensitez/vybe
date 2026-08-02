<?php
// vybe-test: php/php83_features/json_validate_basic
// origin: languages/php/tests/php/test_php83_features.rs

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

$valid = '{"key":"value"}';
$invalid = '{bad json}';
if (function_exists('json_validate')) {
    echo json_validate($valid) ? 'ok' : 'fail';
    echo json_validate($invalid) ? 'ok' : 'fail';
} else {
    json_decode($valid); $ok1 = json_last_error() === JSON_ERROR_NONE;
    json_decode($invalid); $ok2 = json_last_error() === JSON_ERROR_NONE;
    echo $ok1 ? 'ok' : 'fail';
    echo $ok2 ? 'ok' : 'fail';
}

__vybe_check(ob_get_clean(), "okfail");
