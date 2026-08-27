<?php
// vybe-test: php/array_key_destructuring/list_with_reference_modifies_source
// origin: languages/php/tests/php/test_array_key_destructuring.rs

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

echo "list_with_reference_modifies_source_ok";

__vybe_check(ob_get_clean(), "list_with_reference_modifies_source_ok");
