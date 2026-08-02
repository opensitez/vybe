<?php
// vybe-test: php/datetime_immutable/datetime_immutable_diff_invert_flag
// origin: languages/php/tests/php/test_datetime_immutable.rs

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
$a = new DateTimeImmutable('2024-01-10', new DateTimeZone('UTC'));
$b = new DateTimeImmutable('2024-01-01', new DateTimeZone('UTC'));
echo $a->diff($b)->invert;

__vybe_check(ob_get_clean(), "1");
