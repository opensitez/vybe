<?php
// vybe-test: php/datetime_parsing/datetime_create_from_format_with_timezone_object
// origin: languages/php/tests/php/test_datetime_parsing.rs

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

$dt = DateTime::createFromFormat(
    'Y-m-d H:i:s',
    '2024-03-01 10:00:00',
    new DateTimeZone('Asia/Tokyo')
);
echo $dt !== false ? $dt->getTimezone()->getName() : 'bad';

__vybe_check(ob_get_clean(), "Asia/Tokyo");
