<?php
// vybe-test: php/php_datetime_create_from_interface/test_datetime_create_from_interface_mutable_to_immutable
// origin: languages/php/tests/php/test_php_datetime_create_from_interface.rs

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

$mut = new DateTime('2024-12-31 23:59:59', new DateTimeZone('UTC'));
$imm = DateTimeImmutable::createFromInterface($mut);
echo $imm->format('Y-m-d H:i:s') . ':' . get_class($imm), "\n";

__vybe_check(ob_get_clean(), "2024-12-31 23:59:59:DateTimeImmutable");
