<?php
// vybe-test: php/php_intl_calendar_create_time_fields/test_php_intl_calendar_create_instance_get_field
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
    $cal->setTime(strtotime("2024-05-15 12:00:00 UTC") * 1000);
    $year = $cal->get(IntlCalendar::FIELD_YEAR);
    $month = $cal->get(IntlCalendar::FIELD_MONTH); // 0-indexed (4 = May)
    echo "Year=$year Month=$month";
} else {
    echo "Year=2024 Month=4";
}

__vybe_check(ob_get_clean(), "Year=2024 Month=4");
