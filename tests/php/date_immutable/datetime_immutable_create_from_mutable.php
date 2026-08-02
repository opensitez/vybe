<?php
// vybe-test: php/date_immutable/datetime_immutable_create_from_mutable
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

$mutable = new DateTime('2024-11-20 09:00:00');
$immutable = DateTimeImmutable::createFromMutable($mutable);
echo ($immutable instanceof DateTimeImmutable) ? 'yes' : 'no';
echo ':' . $immutable->format('Y-m-d H:i:s');

__vybe_check(ob_get_clean(), "yes:2024-11-20 09:00:00");
