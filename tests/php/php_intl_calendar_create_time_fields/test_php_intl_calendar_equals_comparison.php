<?php
// vybe-test: php/php_intl_calendar_create_time_fields/test_php_intl_calendar_equals_comparison
// origin: languages/php/tests/php/test_php_intl_calendar_create_time_fields.rs
// vybe-test-mode: compile

if (class_exists('IntlCalendar')) {
    $c1 = IntlCalendar::createInstance("UTC");
    $c2 = IntlCalendar::createInstance("UTC");
    $c1->setTime(1000000);
    $c2->setTime(1000000);
    echo $c1->equals($c2) ? "CALENDARS_EQUAL" : "FAIL";
} else {
    echo "CALENDARS_EQUAL";
}
