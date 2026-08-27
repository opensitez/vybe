<?php
// vybe-test: php/first_class_callables/memoize_via_first_class_callable
// origin: languages/php/tests/php/test_first_class_callables.rs

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

echo "memoize_via_first_class_callable_ok";

__vybe_check(ob_get_clean(), "memoize_via_first_class_callable_ok");
