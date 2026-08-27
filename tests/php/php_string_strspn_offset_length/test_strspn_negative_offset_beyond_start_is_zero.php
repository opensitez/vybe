<?php
// vybe-test: php/php_string_strspn_offset_length/test_strspn_negative_offset_beyond_start_is_zero
// origin: languages/php/tests/php/test_php_string_strspn_offset_length.rs

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

echo "test_strspn_negative_offset_beyond_start_is_zero_ok";

__vybe_check(ob_get_clean(), "test_strspn_negative_offset_beyond_start_is_zero_ok");
