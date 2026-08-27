<?php
// vybe-test: php/literals/test_php_octal_and_hex_string_key_edge_cases
// origin: languages/php/tests/php/test_literals.rs

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

echo "test_php_octal_and_hex_string_key_edge_cases_ok";

__vybe_check(ob_get_clean(), "test_php_octal_and_hex_string_key_edge_cases_ok");
