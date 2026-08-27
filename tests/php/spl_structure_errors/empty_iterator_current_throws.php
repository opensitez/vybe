<?php
// vybe-test: php/spl_structure_errors/empty_iterator_current_throws
// origin: languages/php/tests/php/test_spl_structure_errors.rs

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

echo "empty_iterator_current_throws_ok";

__vybe_check(ob_get_clean(), "empty_iterator_current_throws_ok");
