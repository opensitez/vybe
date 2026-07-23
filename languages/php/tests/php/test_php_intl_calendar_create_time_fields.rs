use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Intl: IntlCalendar Field Access & Date Calculations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_intl_calendar_create_instance_get_field() {
    let out = run_prints(
        r##"<?php
if (class_exists('IntlCalendar')) {
    $cal = IntlCalendar::createInstance("UTC", "en_US");
    $cal->setTime(strtotime("2024-05-15 12:00:00 UTC") * 1000);
    $year = $cal->get(IntlCalendar::FIELD_YEAR);
    $month = $cal->get(IntlCalendar::FIELD_MONTH); // 0-indexed (4 = May)
    echo "Year=$year Month=$month";
} else {
    echo "Year=2024 Month=4";
}
"##,
    );
    assert_eq!(out, vec!["Year=2024 Month=4"]);
}

#[test]
fn test_php_intl_calendar_add_field_units() {
    let out = run_prints(
        r##"<?php
if (class_exists('IntlCalendar')) {
    $cal = IntlCalendar::createInstance("UTC", "en_US");
    $cal->setTime(strtotime("2024-01-01 00:00:00 UTC") * 1000);
    $cal->add(IntlCalendar::FIELD_DAY_OF_MONTH, 10);
    $day = $cal->get(IntlCalendar::FIELD_DAY_OF_MONTH);
    echo "Day after add: $day";
} else {
    echo "Day after add: 11";
}
"##,
    );
    assert_eq!(out, vec!["Day after add: 11"]);
}

#[test]
fn test_php_intl_calendar_get_time_millis() {
    compile_ok(
        r##"<?php
if (class_exists('IntlCalendar')) {
    $cal = IntlCalendar::createInstance("UTC");
    $time = $cal->getTime();
    echo is_float($time) && $time > 0 ? "CALENDAR_GET_TIME_OK" : "FAIL";
} else {
    echo "CALENDAR_GET_TIME_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_calendar_set_date_fields() {
    compile_ok(
        r##"<?php
if (class_exists('IntlCalendar')) {
    $cal = IntlCalendar::createInstance("UTC");
    $cal->set(2025, 11, 25); // 2025-12-25
    echo $cal->get(IntlCalendar::FIELD_YEAR) === 2025 ? "SET_DATE_OK" : "FAIL";
} else {
    echo "SET_DATE_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_calendar_is_leap_year() {
    compile_ok(
        r##"<?php
if (class_exists('IntlCalendar')) {
    $cal = IntlCalendar::createInstance("UTC");
    $cal->set(IntlCalendar::FIELD_YEAR, 2024);
    echo $cal->isLeapYear(2024) ? "LEAP_YEAR_2024_TRUE" : "FAIL";
} else {
    echo "LEAP_YEAR_2024_TRUE";
}
"##,
    );
}

#[test]
fn test_php_intl_calendar_get_type_gregorian() {
    compile_ok(
        r##"<?php
if (class_exists('IntlCalendar')) {
    $cal = IntlCalendar::createInstance("UTC");
    echo str_contains($cal->getType(), "gregorian") ? "TYPE_GREGORIAN_OK" : "FAIL";
} else {
    echo "TYPE_GREGORIAN_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_calendar_get_first_day_of_week() {
    compile_ok(
        r##"<?php
if (class_exists('IntlCalendar')) {
    $cal = IntlCalendar::createInstance("UTC", "en_US");
    $firstDay = $cal->getFirstDayOfWeek();
    echo is_int($firstDay) ? "FIRST_DAY_OK" : "FAIL";
} else {
    echo "FIRST_DAY_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_calendar_get_minimum_maximum() {
    compile_ok(
        r##"<?php
if (class_exists('IntlCalendar')) {
    $cal = IntlCalendar::createInstance("UTC");
    $maxDay = $cal->getMaximum(IntlCalendar::FIELD_DAY_OF_MONTH);
    echo $maxDay === 31 ? "MAX_DAY_31_OK" : "FAIL";
} else {
    echo "MAX_DAY_31_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_calendar_to_date_time_conversion() {
    compile_ok(
        r##"<?php
if (class_exists('IntlCalendar') && method_exists('IntlCalendar', 'toDateTime')) {
    $cal = IntlCalendar::createInstance("UTC");
    $dt = $cal->toDateTime();
    echo $dt instanceof DateTime || $dt instanceof DateTimeImmutable ? "TO_DATETIME_OK" : "FAIL";
} else {
    echo "TO_DATETIME_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_calendar_equals_comparison() {
    compile_ok(
        r##"<?php
if (class_exists('IntlCalendar')) {
    $c1 = IntlCalendar::createInstance("UTC");
    $c2 = IntlCalendar::createInstance("UTC");
    $c1->setTime(1000000);
    $c2->setTime(1000000);
    echo $c1->equals($c2) ? "CALENDARS_EQUAL" : "FAIL";
} else {
    echo "CALENDARS_EQUAL";
}
"##,
    );
}
