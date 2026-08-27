<?php
// vybe-test: php/attributes/attribute_is_instance_returns_true_for_matching_class
// origin: languages/php/tests/php/test_attributes.rs

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

echo "attribute_is_instance_returns_true_for_matching_class_ok";

__vybe_check(ob_get_clean(), "attribute_is_instance_returns_true_for_matching_class_ok");
