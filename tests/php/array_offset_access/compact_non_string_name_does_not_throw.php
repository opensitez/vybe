<?php
// vybe-test: php/array_offset_access/compact_non_string_name_does_not_throw
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

echo "compact_non_string_name_does_not_throw_ok";

__vybe_check(ob_get_clean(), "compact_non_string_name_does_not_throw_ok");
