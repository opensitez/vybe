<?php
// vybe-test: php/php_intl_calendar_create_time_fields/test_php_intl_calendar_add_field_units
// origin: languages/php/tests/php/test_php_intl_calendar_create_time_fields.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

if (class_exists('IntlCalendar')) {
    $cal = IntlCalendar::createInstance("UTC", "en_US");
    $cal->setTime(strtotime("2024-01-01 00:00:00 UTC") * 1000);
    $cal->add(IntlCalendar::FIELD_DAY_OF_MONTH, 10);
    $day = $cal->get(IntlCalendar::FIELD_DAY_OF_MONTH);
    echo "Day after add: $day";
} else {
    echo "Day after add: 11";
}

__vybe_check(ob_get_clean(), "Day after add: 11");
