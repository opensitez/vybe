<?php
// vybe-test: php/php84_datetime_create_from_timestamp/test_php84_datetime_create_from_timestamp_integer
// origin: languages/php/tests/php/test_php84_datetime_create_from_timestamp.rs

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

$ts = 1704067200; // 2024-01-01 00:00:00 UTC
if (method_exists('DateTimeImmutable', 'createFromTimestamp')) {
    $dt = DateTimeImmutable::createFromTimestamp($ts);
    echo $dt->format("Y-m-d H:i:s");
} else {
    $dt = (new DateTimeImmutable())->setTimestamp($ts);
    echo $dt->format("Y-m-d H:i:s");
}

__vybe_check(ob_get_clean(), "2024-01-01 00:00:00");
