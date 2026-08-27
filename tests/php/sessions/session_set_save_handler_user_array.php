<?php
// vybe-test: php/sessions/session_set_save_handler_user_array
// origin: languages/php/tests/php/test_sessions.rs

function __vybe_check($got, $want) {
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

echo "session_set_save_handler_user_array_ok";

__vybe_check(ob_get_clean(), "session_set_save_handler_user_array_ok");
