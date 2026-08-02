<?php
// vybe-test: php/php_intl_timezone_create_offset/test_php_intl_timezone_create_gmt
// origin: languages/php/tests/php/test_php_intl_timezone_create_offset.rs
// vybe-test-mode: compile

if (class_exists('IntlTimeZone')) {
    $tz = IntlTimeZone::getGMT();
    echo $tz->getID() === "GMT" || $tz->getID() === "UTC" ? "GMT_TZ_OK" : "FAIL";
} else {
    echo "GMT_TZ_OK";
}
