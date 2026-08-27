<?php
// vybe-test: php/datetime_immutable/datetime_immutable_original_unchanged_after_ops
// origin: languages/php/tests/php/test_datetime_immutable.rs

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

echo "datetime_immutable_original_unchanged_after_ops_ok";

__vybe_check(ob_get_clean(), "datetime_immutable_original_unchanged_after_ops_ok");
