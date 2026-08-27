<?php
// vybe-test: php/php_closures_bind_bindto_scope/test_php_closure_bindto_null_this_unbinding
// origin: languages/php/tests/php/test_php_closures_bind_bindto_scope.rs

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

echo "test_php_closure_bindto_null_this_unbinding_ok";

__vybe_check(ob_get_clean(), "test_php_closure_bindto_null_this_unbinding_ok");
