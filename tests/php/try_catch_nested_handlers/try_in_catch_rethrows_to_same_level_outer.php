<?php
// vybe-test: php/try_catch_nested_handlers/try_in_catch_rethrows_to_same_level_outer
// origin: languages/php/tests/php/test_try_catch_nested_handlers.rs

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

echo "try_in_catch_rethrows_to_same_level_outer_ok";

__vybe_check(ob_get_clean(), "try_in_catch_rethrows_to_same_level_outer_ok");
