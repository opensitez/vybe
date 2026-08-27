<?php
// vybe-test: php/datetime_modify_relative_formats/datetime_modify_relative_midyear_keyword
// origin: languages/php/tests/php/test_datetime_modify_relative_formats.rs

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

echo "datetime_modify_relative_midyear_keyword_ok";

__vybe_check(ob_get_clean(), "datetime_modify_relative_midyear_keyword_ok");
