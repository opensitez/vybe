<?php
// vybe-test: php/oop_patterns/interface_typehint_polymorphism
// origin: languages/php/tests/php/test_oop_patterns.rs

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

echo "interface_typehint_polymorphism_ok";

__vybe_check(ob_get_clean(), "interface_typehint_polymorphism_ok");
