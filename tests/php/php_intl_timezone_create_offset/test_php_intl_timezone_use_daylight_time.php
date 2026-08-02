<?php
// vybe-test: php/php_intl_timezone_create_offset/test_php_intl_timezone_use_daylight_time
// origin: languages/php/tests/php/test_php_intl_timezone_create_offset.rs
// vybe-test-mode: compile

if (class_exists('IntlTimeZone')) {
    $tz = IntlTimeZone::createTimeZone("Europe/Paris");
    echo $tz->useDaylightTime() ? "DST_USED_PARIS" : "FAIL";
} else {
    echo "DST_USED_PARIS";
}
