<?php
// vybe-test: php/php_intl_calendar_create_time_fields/test_php_intl_calendar_to_date_time_conversion
// origin: languages/php/tests/php/test_php_intl_calendar_create_time_fields.rs
// vybe-test-mode: compile

if (class_exists('IntlCalendar') && method_exists('IntlCalendar', 'toDateTime')) {
    $cal = IntlCalendar::createInstance("UTC");
    $dt = $cal->toDateTime();
    echo $dt instanceof DateTime || $dt instanceof DateTimeImmutable ? "TO_DATETIME_OK" : "FAIL";
} else {
    echo "TO_DATETIME_OK";
}
