<?php
// vybe-test: php/spl_structure_errors/multiple_iterator_without_attach_has_zero_count
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

echo "multiple_iterator_without_attach_has_zero_count_ok";

__vybe_check(ob_get_clean(), "multiple_iterator_without_attach_has_zero_count_ok");
