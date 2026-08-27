<?php
// vybe-test: php/php_json_decode_object_as_array_flag/test_json_decode_object_as_array_flag
// origin: languages/php/tests/php/test_php_json_decode_object_as_array_flag.rs

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

echo "test_json_decode_object_as_array_flag_ok";

__vybe_check(ob_get_clean(), "test_json_decode_object_as_array_flag_ok");
