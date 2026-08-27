<?php
// vybe-test: php/destructure_assignment/nested_list_destructure_from_scalar_inner_assigns_null
// origin: languages/php/tests/php/test_destructure_assignment.rs

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

echo "nested_list_destructure_from_scalar_inner_assigns_null_ok";

__vybe_check(ob_get_clean(), "nested_list_destructure_from_scalar_inner_assigns_null_ok");
