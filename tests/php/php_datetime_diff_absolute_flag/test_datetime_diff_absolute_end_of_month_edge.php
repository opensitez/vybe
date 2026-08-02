<?php
// vybe-test: php/php_datetime_diff_absolute_flag/test_datetime_diff_absolute_end_of_month_edge
// origin: languages/php/tests/php/test_php_datetime_diff_absolute_flag.rs

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

$d1 = new DateTime('2024-01-31');
$d2 = new DateTime('2024-02-29');
$diff = $d1->diff($d2, true);
echo $diff->days . ':' . $diff->m . ':' . $diff->d;

__vybe_check(ob_get_clean(), "29:0:29");
