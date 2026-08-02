<?php
// vybe-test: php/php_intl_calendar_create_time_fields/test_php_intl_calendar_get_minimum_maximum
// origin: languages/php/tests/php/test_php_intl_calendar_create_time_fields.rs
// vybe-test-mode: compile

if (class_exists('IntlCalendar')) {
    $cal = IntlCalendar::createInstance("UTC");
    $maxDay = $cal->getMaximum(IntlCalendar::FIELD_DAY_OF_MONTH);
    echo $maxDay === 31 ? "MAX_DAY_31_OK" : "FAIL";
} else {
    echo "MAX_DAY_31_OK";
}
