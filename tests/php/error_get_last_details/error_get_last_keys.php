<?php
// vybe-test: php/error_get_last_details/error_get_last_keys
// origin: languages/php/tests/php/test_error_get_last_details.rs

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

@trigger_error("custom msg", E_USER_NOTICE);
$err = error_get_last();
if ($err) {
    echo $err['type'] === E_USER_NOTICE ? "notice|" : "fail|";
    echo $err['message'] === "custom msg" ? "msg" : "fail";
}

__vybe_check(ob_get_clean(), "notice|msg");
