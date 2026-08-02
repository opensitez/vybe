<?php
// vybe-test: php/php_datetime_diff_format_timezone/test_php_datetime_diff_with_timezone_objects
// origin: languages/php/tests/php/test_php_datetime_diff_format_timezone.rs

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

$utc = new DateTimeImmutable('2024-01-01 00:00:00', new DateTimeZone('UTC'));
$tokyo = new DateTimeImmutable('2024-01-01 00:00:00', new DateTimeZone('Asia/Tokyo'));
$diff = $utc->diff($tokyo, true);
echo $diff->invert ? 'inv' : 'pos';
echo '|';
echo $diff->days;

__vybe_check(ob_get_clean(), "pos|0");
