<?php
// vybe-test: php/datetime_parsing/datetime_immutable_create_from_format_bad_input
// origin: languages/php/tests/php/test_datetime_parsing.rs

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

echo "datetime_immutable_create_from_format_bad_input_ok";

__vybe_check(ob_get_clean(), "datetime_immutable_create_from_format_bad_input_ok");
