<?php
// vybe-test: php/catch_type_union_order/error_subclass_not_caught_by_exception_only_handler
// origin: languages/php/tests/php/test_catch_type_union_order.rs

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

echo "error_subclass_not_caught_by_exception_only_handler_ok";

__vybe_check(ob_get_clean(), "error_subclass_not_caught_by_exception_only_handler_ok");
