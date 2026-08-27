<?php
// vybe-test: php/builtins/string_post_inc_date_loop_terminates
// origin: languages/php/tests/php/test_builtins.rs

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

echo "string_post_inc_date_loop_terminates_ok";

__vybe_check(ob_get_clean(), "string_post_inc_date_loop_terminates_ok");
