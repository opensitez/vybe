<?php
// vybe-test: php/trigger_error_user_levels/trigger_error_levels
// origin: languages/php/tests/php/test_trigger_error_user_levels.rs

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

set_error_handler(function($errno, $errstr) {
    echo "$errno:$errstr|";
    return true;
});

trigger_error("warn", E_USER_WARNING);
trigger_error("notice", E_USER_NOTICE);
trigger_error("deprecated", E_USER_DEPRECATED);

__vybe_check(ob_get_clean(), "512:warn|1024:notice|16384:deprecated|");
