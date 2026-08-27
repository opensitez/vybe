<?php
// vybe-test: php/functional_patterns/trampoline_tail_recursion
// origin: languages/php/tests/php/test_functional_patterns.rs

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

echo "trampoline_tail_recursion_ok";

__vybe_check(ob_get_clean(), "trampoline_tail_recursion_ok");
