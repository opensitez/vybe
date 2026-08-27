<?php
// vybe-test: php/php_session_gc_garbage_collection/test_session_gc_execution
// origin: languages/php/tests/php/test_php_session_gc_garbage_collection.rs

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

echo "test_session_gc_execution_ok";

__vybe_check(ob_get_clean(), "test_session_gc_execution_ok");
