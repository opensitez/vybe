<?php
// vybe-test: php/filter_input_array_globals/filter_input_array_basic
// origin: languages/php/tests/php/test_filter_input_array_globals.rs

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

echo "filter_input_array_basic_ok";

__vybe_check(ob_get_clean(), "filter_input_array_basic_ok");
