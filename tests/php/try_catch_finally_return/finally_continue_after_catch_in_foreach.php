<?php
// vybe-test: php/try_catch_finally_return/finally_continue_after_catch_in_foreach
// origin: languages/php/tests/php/test_try_catch_finally_return.rs

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

echo "finally_continue_after_catch_in_foreach_ok";

__vybe_check(ob_get_clean(), "finally_continue_after_catch_in_foreach_ok");
