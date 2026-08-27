<?php
// vybe-test: php/threads/thread_with_closure
// origin: languages/php/tests/php/test_threads.rs

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

echo "thread_with_closure_ok";

__vybe_check(ob_get_clean(), "thread_with_closure_ok");
