<?php
// vybe-test: php/php_datetime_timezones/php_datetime_timezone_construct_invalid_zone_errors
// origin: languages/php/tests/php/test_php_datetime_timezones.rs

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

try {
    new DateTimeZone('Not/AZone');
    echo 'ok';
} catch (Throwable $e) {
    echo 'err';
}

__vybe_check(ob_get_clean(), "err");
