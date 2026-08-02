<?php
// vybe-test: php/datetime_diff_absolute/datetime_diff_absolute_year_month_parts
// origin: languages/php/tests/php/test_datetime_diff_absolute.rs

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

$dt1 = new DateTime('2023-01-15');
$dt2 = new DateTime('2024-03-20');
$diff = $dt1->diff($dt2, true);
echo $diff->y . "," . $diff->m . "," . $diff->d;

__vybe_check(ob_get_clean(), "1,2,5");
