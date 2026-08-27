<?php
// vybe-test: php/php_spl_observer_subject_pattern/test_php_spl_object_storage_contains_and_offset
// origin: languages/php/tests/php/test_php_spl_observer_subject_pattern.rs

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

echo "test_php_spl_object_storage_contains_and_offset_ok";

__vybe_check(ob_get_clean(), "test_php_spl_object_storage_contains_and_offset_ok");
