<?php
// vybe-test: php/datetime/datetime_set_timezone_with_copy
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

$dt = new DateTime('2024-01-02 12:00:00', new DateTimeZone('UTC'));
$other = clone $dt;
$other->setTimezone(new DateTimeZone('Asia/Tokyo'));
echo $dt->format('H:i');
echo '|';
echo $other->format('H:i');

__vybe_check(ob_get_clean(), "12:00|21:00");
