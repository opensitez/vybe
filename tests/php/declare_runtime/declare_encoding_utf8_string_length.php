<?php
// vybe-test: php/declare_runtime/declare_encoding_utf8_string_length
// origin: languages/php/tests/php/test_declare_runtime.rs

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

echo "declare_encoding_utf8_string_length_ok";

__vybe_check(ob_get_clean(), "declare_encoding_utf8_string_length_ok");
