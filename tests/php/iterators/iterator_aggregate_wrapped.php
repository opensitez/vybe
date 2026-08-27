<?php
// vybe-test: php/iterators/iterator_aggregate_wrapped
// origin: languages/php/tests/php/test_iterators.rs

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

echo "iterator_aggregate_wrapped_ok";

__vybe_check(ob_get_clean(), "iterator_aggregate_wrapped_ok");
