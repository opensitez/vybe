<?php
// vybe-test: php/php_datetime_period_iso_specifiers/test_date_period_iso_get_start_end_dates
// origin: languages/php/tests/php/test_php_datetime_period_iso_specifiers.rs

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

$period = new DatePeriod('R2/2024-12-30T00:00:00Z/P1D');
$end = $period->getEndDate();
echo $period->getStartDate()->format('Y-m-d');
echo '|';
echo $end instanceof DateTimeInterface ? $end->format('Y-m-d') : 'none';

__vybe_check(ob_get_clean(), "2024-12-30|2024-12-31");
