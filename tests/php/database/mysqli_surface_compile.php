<?php
// vybe-test: php/database/mysqli_surface_compile
// origin: languages/php/tests/php/test_database.rs

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

echo "mysqli_surface_compile_ok";

__vybe_check(ob_get_clean(), "mysqli_surface_compile_ok");
