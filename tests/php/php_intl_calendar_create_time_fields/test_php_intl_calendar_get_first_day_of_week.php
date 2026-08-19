<?php
// vybe-test: php/php_intl_calendar_create_time_fields/test_php_intl_calendar_get_first_day_of_week
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
    $firstDay = $cal->getFirstDayOfWeek();
    echo is_int($firstDay) ? "FIRST_DAY_OK" : "FAIL";
} else {
    echo "FIRST_DAY_OK";
}


__vybe_check(ob_get_clean(), "FIRST_DAY_OK");
