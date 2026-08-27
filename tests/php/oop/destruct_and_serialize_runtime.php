<?php
// vybe-test: php/oop/destruct_and_serialize_runtime
// origin: languages/php/tests/php/test_oop.rs

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

echo "destruct_and_serialize_runtime_ok";

__vybe_check(ob_get_clean(), "destruct_and_serialize_runtime_ok");
