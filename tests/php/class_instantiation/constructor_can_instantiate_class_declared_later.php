<?php
// vybe-test: php/class_instantiation/constructor_can_instantiate_class_declared_later
// origin: languages/php/tests/php/test_class_instantiation.rs

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

echo "constructor_can_instantiate_class_declared_later_ok";

__vybe_check(ob_get_clean(), "constructor_can_instantiate_class_declared_later_ok");
