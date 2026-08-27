<?php
// vybe-test: php/serialize/unserialize_false_on_garbage
// origin: languages/php/tests/php/test_serialize.rs

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

echo "unserialize_false_on_garbage_ok";

__vybe_check(ob_get_clean(), "unserialize_false_on_garbage_ok");
