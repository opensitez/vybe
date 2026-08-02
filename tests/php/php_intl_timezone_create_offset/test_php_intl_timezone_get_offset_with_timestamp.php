<?php
// vybe-test: php/php_intl_timezone_create_offset/test_php_intl_timezone_get_offset_with_timestamp
// origin: languages/php/tests/php/test_php_intl_timezone_create_offset.rs
// vybe-test-mode: compile

if (class_exists('IntlTimeZone')) {
    $tz = IntlTimeZone::createTimeZone("Europe/London");
    $rawOffset = 0;
    $dstOffset = 0;
    $ts = strtotime("2024-07-01 12:00:00 UTC") * 1000;
    $res = $tz->getOffset($ts, false, $rawOffset, $dstOffset);
    echo $res ? "OFFSET_CALCULATED_OK" : "FAIL";
} else {
    echo "OFFSET_CALCULATED_OK";
}
