<?php
// vybe-test: php/php_dateperiod_include_end_date_flag/test_date_period_start_utc_timezone_option
// origin: languages/php/tests/php/test_php_dateperiod_include_end_date_flag.rs

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

$tz = new DateTimeZone('Europe/London');
$start = new DateTime('2024-03-01 00:00:00', $tz);
$end = new DateTime('2024-03-05 00:00:00', $tz);
$period = new DatePeriod($start, new DateInterval('P1D'), $end, DatePeriod::INCLUDE_END_DATE);
$date = [];
foreach ($period as $dt) {
    $date[] = $dt->getTimezone()->getName();
}
echo implode(',', array_unique($date));

__vybe_check(ob_get_clean(), "Europe/London");
