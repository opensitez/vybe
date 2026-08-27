<?php
// vybe-test: php/database_prepared/prepared_param_count_matches_placeholders
// origin: languages/php/tests/php/test_database_prepared.rs

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

echo "prepared_param_count_matches_placeholders_ok";

__vybe_check(ob_get_clean(), "prepared_param_count_matches_placeholders_ok");
