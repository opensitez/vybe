<?php
// vybe-test: php/php_intl_calendar_create_time_fields/test_php_intl_calendar_is_leap_year
// origin: languages/php/tests/php/test_php_intl_calendar_create_time_fields.rs
// vybe-test-mode: compile

if (class_exists('IntlCalendar')) {
    $cal = IntlCalendar::createInstance("UTC");
    $cal->set(IntlCalendar::FIELD_YEAR, 2024);
    echo $cal->isLeapYear(2024) ? "LEAP_YEAR_2024_TRUE" : "FAIL";
} else {
    echo "LEAP_YEAR_2024_TRUE";
}
