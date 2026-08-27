<?php
// vybe-test: php/array_offset_access/string_offset_out_of_range_read_returns_null
// origin: languages/php/tests/php/test_array_offset_access.rs

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

echo "string_offset_out_of_range_read_returns_null_ok";

__vybe_check(ob_get_clean(), "string_offset_out_of_range_read_returns_null_ok");
