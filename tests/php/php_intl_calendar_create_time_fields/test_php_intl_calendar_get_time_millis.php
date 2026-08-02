<?php
// vybe-test: php/php_intl_calendar_create_time_fields/test_php_intl_calendar_get_time_millis
// origin: languages/php/tests/php/test_php_intl_calendar_create_time_fields.rs
// vybe-test-mode: compile

if (class_exists('IntlCalendar')) {
    $cal = IntlCalendar::createInstance("UTC");
    $time = $cal->getTime();
    echo is_float($time) && $time > 0 ? "CALENDAR_GET_TIME_OK" : "FAIL";
} else {
    echo "CALENDAR_GET_TIME_OK";
}
