<?php
// vybe-test: php/idate_timestamp_parts/idate_month_day_names_runtime
// origin: languages/php/tests/php/test_idate_timestamp_parts.rs

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

date_default_timezone_set('UTC');
$timestamp = mktime(0, 0, 0, 12, 31, 2023);
echo idate('L', $timestamp) . "|" . idate('t', $timestamp);

__vybe_check(ob_get_clean(), "1|31");
