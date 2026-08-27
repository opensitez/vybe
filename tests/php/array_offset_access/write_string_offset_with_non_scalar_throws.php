<?php
// vybe-test: php/array_offset_access/write_string_offset_with_non_scalar_throws
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

echo "write_string_offset_with_non_scalar_throws_ok";

__vybe_check(ob_get_clean(), "write_string_offset_with_non_scalar_throws_ok");
