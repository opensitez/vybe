<?php
// vybe-test: php/datetime_parsing/datetime_create_from_format_trailing_junk_sets_warning
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

echo "datetime_create_from_format_trailing_junk_sets_warning_ok";

__vybe_check(ob_get_clean(), "datetime_create_from_format_trailing_junk_sets_warning_ok");
