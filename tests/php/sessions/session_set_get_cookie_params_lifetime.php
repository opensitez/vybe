<?php
// vybe-test: php/sessions/session_set_get_cookie_params_lifetime
// origin: languages/php/tests/php/test_sessions.rs

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

session_set_cookie_params(['lifetime' => 3600]);
session_start();
$p = session_get_cookie_params();
echo $p['lifetime'];

__vybe_check(ob_get_clean(), "3600");
