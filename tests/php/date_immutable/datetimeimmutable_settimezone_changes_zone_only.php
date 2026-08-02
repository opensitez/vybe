<?php
// vybe-test: php/date_immutable/datetimeimmutable_settimezone_changes_zone_only
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

$d = new DateTimeImmutable('2024-06-15 12:00:00', new DateTimeZone('UTC'));
$d2 = $d->setTimezone(new DateTimeZone('America/Los_Angeles'));
echo $d->getTimezone()->getName() . ':' . $d2->getTimezone()->getName();

__vybe_check(ob_get_clean(), "UTC:America/Los_Angeles");
