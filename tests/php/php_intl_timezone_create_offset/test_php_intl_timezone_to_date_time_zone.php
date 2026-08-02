<?php
// vybe-test: php/php_intl_timezone_create_offset/test_php_intl_timezone_to_date_time_zone
// origin: languages/php/tests/php/test_php_intl_timezone_create_offset.rs
// vybe-test-mode: compile

if (class_exists('IntlTimeZone') && method_exists('IntlTimeZone', 'toDateTimeZone')) {
    $tz = IntlTimeZone::createTimeZone("Asia/Tokyo");
    $dtz = $tz->toDateTimeZone();
    echo $dtz instanceof DateTimeZone ? "TO_DATETIMEZONE_OK" : "FAIL";
} else {
    echo "TO_DATETIMEZONE_OK";
}
