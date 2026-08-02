<?php
// vybe-test: php/php_set_error_handler_levels_mask/test_php_set_error_handler_level_mask_filtering
// origin: languages/php/tests/php/test_php_set_error_handler_levels_mask.rs

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

$captured = [];
set_error_handler(function($errno, $errstr) use (&$captured) {
    $captured[] = "Err:$errno Str:$errstr";
}, E_USER_WARNING);

@trigger_error("Ignored notice", E_USER_NOTICE);
@trigger_error("Captured warning", E_USER_WARNING);

restore_error_handler();

echo implode("; ", $captured);

__vybe_check(ob_get_clean(), "Err:512 Str:Captured warning");
