<?php
// vybe-test: php/php_array_key_operations_indexing/test_php_array_key_casting_edge_cases
// origin: languages/php/tests/php/test_php_array_key_operations_indexing.rs

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

echo "test_php_array_key_casting_edge_cases_ok";

__vybe_check(ob_get_clean(), "test_php_array_key_casting_edge_cases_ok");
