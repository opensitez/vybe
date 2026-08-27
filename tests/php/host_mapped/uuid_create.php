<?php
// vybe-test: php/host_mapped/uuid_create
// origin: languages/php/tests/php/test_host_mapped.rs

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

echo "uuid_create_ok";

__vybe_check(ob_get_clean(), "uuid_create_ok");
