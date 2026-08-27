<?php
// vybe-test: php/class_instantiation/extend_final_class_at_parse_time_fails
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

echo "extend_final_class_at_parse_time_fails_ok";

__vybe_check(ob_get_clean(), "extend_final_class_at_parse_time_fails_ok");
