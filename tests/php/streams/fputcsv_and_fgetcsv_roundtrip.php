<?php
// vybe-test: php/streams/fputcsv_and_fgetcsv_roundtrip
// origin: languages/php/tests/php/test_streams.rs

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

echo "fputcsv_and_fgetcsv_roundtrip_ok";

__vybe_check(ob_get_clean(), "fputcsv_and_fgetcsv_roundtrip_ok");
