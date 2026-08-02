<?php
// vybe-test: php/php_intl_calendar_create_time_fields/test_php_intl_calendar_get_first_day_of_week
// origin: languages/php/tests/php/test_php_intl_calendar_create_time_fields.rs
// vybe-test-mode: compile

if (class_exists('IntlCalendar')) {
    $cal = IntlCalendar::createInstance("UTC", "en_US");
    $firstDay = $cal->getFirstDayOfWeek();
    echo is_int($firstDay) ? "FIRST_DAY_OK" : "FAIL";
} else {
    echo "FIRST_DAY_OK";
}
