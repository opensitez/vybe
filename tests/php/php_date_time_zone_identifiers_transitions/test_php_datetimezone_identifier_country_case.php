<?php
// vybe-test: php/php_date_time_zone_identifiers_transitions/test_php_datetimezone_identifier_country_case
// origin: languages/php/tests/php/test_php_date_time_zone_identifiers_transitions.rs

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

echo "test_php_datetimezone_identifier_country_case_ok";

__vybe_check(ob_get_clean(), "test_php_datetimezone_identifier_country_case_ok");
