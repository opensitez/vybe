<?php
// vybe-test: php/sessions/session_cookie_params_httponly_flag
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

session_set_cookie_params(['httponly' => true]);
$p = session_get_cookie_params();
echo $p['httponly'] ? '1' : '0';

__vybe_check(ob_get_clean(), "1");
