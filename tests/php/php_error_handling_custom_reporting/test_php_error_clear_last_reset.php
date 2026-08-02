<?php
// vybe-test: php/php_error_handling_custom_reporting/test_php_error_clear_last_reset
// origin: languages/php/tests/php/test_php_error_handling_custom_reporting.rs

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

@trigger_error("Test notice", E_USER_NOTICE);
$err1 = error_get_last();
error_clear_last();
$err2 = error_get_last();

echo ($err1 !== null ? "HAD_ERROR" : "NO") . " | " . ($err2 === null ? "CLEARED" : "NOT_CLEARED");

__vybe_check(ob_get_clean(), "HAD_ERROR | CLEARED");
