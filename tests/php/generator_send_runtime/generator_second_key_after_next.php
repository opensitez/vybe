<?php
// vybe-test: php/generator_send_runtime/generator_second_key_after_next
// origin: languages/php/tests/php/test_generator_send_runtime.rs

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

echo "generator_second_key_after_next_ok";

__vybe_check(ob_get_clean(), "generator_second_key_after_next_ok");
