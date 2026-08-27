<?php
// vybe-test: php/sessions/session_gc_probability_returns_int
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

echo "session_gc_probability_returns_int_ok";

__vybe_check(ob_get_clean(), "session_gc_probability_returns_int_ok");
