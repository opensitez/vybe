<?php
// vybe-test: php/datetime/datetime_create_from_format_with_timezone_designator
// origin: languages/php/tests/php/test_datetime.rs

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

$dt = DateTime::createFromFormat('Y-m-d H:i:s P', '2024-08-01 10:15:00 +02:00');
echo $dt !== false ? 'ok' : 'bad';
echo $dt ? '|' . $dt->format('P') : '|err';

__vybe_check(ob_get_clean(), "ok|+02:00");
