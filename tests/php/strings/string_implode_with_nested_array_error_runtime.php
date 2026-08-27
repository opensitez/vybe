<?php
// vybe-test: php/strings/string_implode_with_nested_array_error_runtime
// origin: languages/php/tests/php/test_strings.rs

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

echo "string_implode_with_nested_array_error_runtime_ok";

__vybe_check(ob_get_clean(), "string_implode_with_nested_array_error_runtime_ok");
