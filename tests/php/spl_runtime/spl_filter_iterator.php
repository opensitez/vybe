<?php
// vybe-test: php/spl_runtime/spl_filter_iterator
// origin: languages/php/tests/php/test_spl_runtime.rs

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

echo "spl_filter_iterator_ok";

__vybe_check(ob_get_clean(), "spl_filter_iterator_ok");
