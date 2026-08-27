<?php
// vybe-test: php/advanced_oop/unresolved_trait_method_collision_is_rejected
// origin: languages/php/tests/php/test_advanced_oop.rs

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

echo "unresolved_trait_method_collision_is_rejected_ok";

__vybe_check(ob_get_clean(), "unresolved_trait_method_collision_is_rejected_ok");
