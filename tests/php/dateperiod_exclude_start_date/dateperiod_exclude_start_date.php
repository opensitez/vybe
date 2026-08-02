<?php
// vybe-test: php/dateperiod_exclude_start_date/dateperiod_exclude_start_date
// origin: languages/php/tests/php/test_dateperiod_exclude_start_date.rs

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

$start = new DateTime('2020-01-01');
$end = new DateTime('2020-01-04');
$interval = new DateInterval('P1D');

$period = new DatePeriod($start, $interval, $end, DatePeriod::EXCLUDE_START_DATE);
$out = [];
foreach ($period as $dt) {
    $out[] = $dt->format('Y-m-d');
}
echo implode(',', $out);

__vybe_check(ob_get_clean(), "2020-01-02,2020-01-03");
