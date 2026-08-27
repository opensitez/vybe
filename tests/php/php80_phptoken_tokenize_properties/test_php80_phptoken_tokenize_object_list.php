<?php
// vybe-test: php/php80_phptoken_tokenize_properties/test_php80_phptoken_tokenize_object_list
// origin: languages/php/tests/php/test_php80_phptoken_tokenize_properties.rs

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

echo "test_php80_phptoken_tokenize_object_list_ok";

__vybe_check(ob_get_clean(), "test_php80_phptoken_tokenize_object_list_ok");
