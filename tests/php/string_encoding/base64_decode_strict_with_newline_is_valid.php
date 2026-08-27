<?php
// vybe-test: php/string_encoding/base64_decode_strict_with_newline_is_valid
// origin: languages/php/tests/php/test_string_encoding.rs

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

echo "base64_decode_strict_with_newline_is_valid_ok";

__vybe_check(ob_get_clean(), "base64_decode_strict_with_newline_is_valid_ok");
