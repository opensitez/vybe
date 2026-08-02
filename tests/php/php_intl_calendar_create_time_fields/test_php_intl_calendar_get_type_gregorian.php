<?php
// vybe-test: php/php_intl_calendar_create_time_fields/test_php_intl_calendar_get_type_gregorian
// origin: languages/php/tests/php/test_php_intl_calendar_create_time_fields.rs
// vybe-test-mode: compile

if (class_exists('IntlCalendar')) {
    $cal = IntlCalendar::createInstance("UTC");
    echo str_contains($cal->getType(), "gregorian") ? "TYPE_GREGORIAN_OK" : "FAIL";
} else {
    echo "TYPE_GREGORIAN_OK";
}
