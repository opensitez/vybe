<?php
// vybe-test: php/php_datetime_create_from_format_errors/test_datetime_create_from_format_timezone_alias
// origin: languages/php/tests/php/test_php_datetime_create_from_format_errors.rs

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

$dt = DateTime::createFromFormat('Y-m-d H:i:s P', '2024-03-01 12:00:00 -0500');
echo $dt !== false ? $dt->format('P') : 'bad';

__vybe_check(ob_get_clean(), "-05:00");
