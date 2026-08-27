<?php
// vybe-test: php/loops/foreach_string_iterates_bytes_in_php_82
// origin: languages/php/tests/php/test_loops.rs

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

echo "foreach_string_iterates_bytes_in_php_82_ok";

__vybe_check(ob_get_clean(), "foreach_string_iterates_bytes_in_php_82_ok");
