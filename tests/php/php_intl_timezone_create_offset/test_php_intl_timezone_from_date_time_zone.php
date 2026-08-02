<?php
// vybe-test: php/php_intl_timezone_create_offset/test_php_intl_timezone_from_date_time_zone
// origin: languages/php/tests/php/test_php_intl_timezone_create_offset.rs
// vybe-test-mode: compile

if (class_exists('IntlTimeZone') && method_exists('IntlTimeZone', 'fromDateTimeZone')) {
    $dtz = new DateTimeZone("UTC");
    $itz = IntlTimeZone::fromDateTimeZone($dtz);
    echo $itz instanceof IntlTimeZone ? "FROM_DATETIMEZONE_OK" : "FAIL";
} else {
    echo "FROM_DATETIMEZONE_OK";
}
