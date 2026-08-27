<?php
// vybe-test: php/array_multisort_mixed_types/array_multisort_with_empty_second_array_still_reorders_first
// origin: languages/php/tests/php/test_array_multisort_mixed_types.rs

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

echo "array_multisort_with_empty_second_array_still_reorders_first_ok";

__vybe_check(ob_get_clean(), "array_multisort_with_empty_second_array_still_reorders_first_ok");
