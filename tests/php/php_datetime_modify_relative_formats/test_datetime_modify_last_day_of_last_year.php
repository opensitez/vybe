<?php
// vybe-test: php/php_datetime_modify_relative_formats/test_datetime_modify_last_day_of_last_year
// origin: languages/php/tests/php/test_php_datetime_modify_relative_formats.rs

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

$dt = new DateTime('2024-07-19', new DateTimeZone('UTC'));
$dt->modify('last day of last year');
echo $dt->format('Y-m-d'), "\n";

__vybe_check(ob_get_clean(), "2023-12-31");
