<?php
// vybe-test: php/php_string_searching_substring_positions/test_php_substr_compare_case_sensitivity_offset
// origin: languages/php/tests/php/test_php_string_searching_substring_positions.rs

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

echo "test_php_substr_compare_case_sensitivity_offset_ok";

__vybe_check(ob_get_clean(), "test_php_substr_compare_case_sensitivity_offset_ok");
