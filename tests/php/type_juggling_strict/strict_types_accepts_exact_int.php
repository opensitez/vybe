<?php
// vybe-test: php/type_juggling_strict/strict_types_accepts_exact_int
// origin: languages/php/tests/php/test_type_juggling_strict.rs

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

echo "strict_types_accepts_exact_int_ok";

__vybe_check(ob_get_clean(), "strict_types_accepts_exact_int_ok");
