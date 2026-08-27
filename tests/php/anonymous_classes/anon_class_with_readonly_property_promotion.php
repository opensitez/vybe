<?php
// vybe-test: php/anonymous_classes/anon_class_with_readonly_property_promotion
// origin: languages/php/tests/php/test_anonymous_classes.rs

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

echo "anon_class_with_readonly_property_promotion_ok";

__vybe_check(ob_get_clean(), "anon_class_with_readonly_property_promotion_ok");
