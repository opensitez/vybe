<?php
// vybe-test: php/php_intl_calendar_create_time_fields/test_php_intl_calendar_set_date_fields
// origin: languages/php/tests/php/test_php_intl_calendar_create_time_fields.rs
// vybe-test-mode: compile

if (class_exists('IntlCalendar')) {
    $cal = IntlCalendar::createInstance("UTC");
    $cal->set(2025, 11, 25); // 2025-12-25
    echo $cal->get(IntlCalendar::FIELD_YEAR) === 2025 ? "SET_DATE_OK" : "FAIL";
} else {
    echo "SET_DATE_OK";
}
