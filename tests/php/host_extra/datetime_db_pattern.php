<?php
// vybe-test: php/host_extra/datetime_db_pattern
// origin: languages/php/tests/php/test_host_extra.rs

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

echo "datetime_db_pattern_ok";

__vybe_check(ob_get_clean(), "datetime_db_pattern_ok");
