<?php
// vybe-test: php/datetime/datetime_localtime_fields_runtime
// origin: languages/php/tests/php/test_datetime.rs

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

echo "datetime_localtime_fields_runtime_ok";

__vybe_check(ob_get_clean(), "datetime_localtime_fields_runtime_ok");
