<?php
// vybe-test: php/php_datetime_immutable_create_from_format/test_datetime_immutable_create_from_format_reset_time
// origin: languages/php/tests/php/test_php_datetime_immutable_create_from_format.rs

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

$dt = DateTimeImmutable::createFromFormat('!d-m-Y', '25-12-2024', new DateTimeZone('UTC'));
echo $dt->format('Y-m-d H:i:s'), "\n";

__vybe_check(ob_get_clean(), "2024-12-25 00:00:00");
