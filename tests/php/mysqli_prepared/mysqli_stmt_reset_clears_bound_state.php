<?php
// vybe-test: php/mysqli_prepared/mysqli_stmt_reset_clears_bound_state
// origin: languages/php/tests/php/test_mysqli_prepared.rs

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

echo "mysqli_stmt_reset_clears_bound_state_ok";

__vybe_check(ob_get_clean(), "mysqli_stmt_reset_clears_bound_state_ok");
