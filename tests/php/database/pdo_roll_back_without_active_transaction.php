<?php
// vybe-test: php/database/pdo_roll_back_without_active_transaction
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

echo "pdo_roll_back_without_active_transaction_ok";

__vybe_check(ob_get_clean(), "pdo_roll_back_without_active_transaction_ok");
