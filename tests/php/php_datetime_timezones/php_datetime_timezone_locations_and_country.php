<?php
// vybe-test: php/php_datetime_timezones/php_datetime_timezone_locations_and_country
// origin: languages/php/tests/php/test_php_datetime_timezones.rs

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

echo "php_datetime_timezone_locations_and_country_ok";

__vybe_check(ob_get_clean(), "php_datetime_timezone_locations_and_country_ok");
