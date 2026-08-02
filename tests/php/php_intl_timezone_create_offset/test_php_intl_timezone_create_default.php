<?php
// vybe-test: php/php_intl_timezone_create_offset/test_php_intl_timezone_create_default
// origin: languages/php/tests/php/test_php_intl_timezone_create_offset.rs
// vybe-test-mode: compile

if (class_exists('IntlTimeZone')) {
    $tz = IntlTimeZone::createDefault();
    echo strlen($tz->getID()) > 0 ? "DEFAULT_TZ_OK" : "FAIL";
} else {
    echo "DEFAULT_TZ_OK";
}
