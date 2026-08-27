<?php
// vybe-test: php/unserialize_errors/unserialize_truncated_object_returns_false
// origin: languages/php/tests/php/test_unserialize_errors.rs

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

echo "unserialize_truncated_object_returns_false_ok";

__vybe_check(ob_get_clean(), "unserialize_truncated_object_returns_false_ok");
