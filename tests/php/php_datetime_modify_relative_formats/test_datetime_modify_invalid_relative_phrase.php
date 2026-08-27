<?php
// vybe-test: php/php_datetime_modify_relative_formats/test_datetime_modify_invalid_relative_phrase
// origin: languages/php/tests/php/test_php_datetime_modify_relative_formats.rs

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

echo "test_datetime_modify_invalid_relative_phrase_ok";

__vybe_check(ob_get_clean(), "test_datetime_modify_invalid_relative_phrase_ok");
