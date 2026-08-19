<?php
// vybe-test: php/date_immutable/datetime_immutable_set_timestamp_preserved
// origin: languages/php/tests/php/test_date_immutable.rs

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

$d = new DateTimeImmutable('2024-01-01', new DateTimeZone('UTC'));
$d2 = $d->setTimestamp(1704067200);
echo $d->format('U') === '1704067200' ? 'old' : 'not old';
echo ':' . $d2->format('Y-m-d');

__vybe_check(ob_get_clean(), "old:2024-01-01");
