<?php
// vybe-test: php/error_clear_last_reset/error_clear_last_basic
// origin: languages/php/tests/php/test_error_clear_last_reset.rs

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

@trigger_error("test", E_USER_WARNING);
error_clear_last();
$err = error_get_last();
echo is_null($err) ? "null" : "not";

__vybe_check(ob_get_clean(), "null");
