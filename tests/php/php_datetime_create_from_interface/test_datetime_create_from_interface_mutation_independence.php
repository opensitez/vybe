<?php
// vybe-test: php/php_datetime_create_from_interface/test_datetime_create_from_interface_mutation_independence
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

$imm = new DateTimeImmutable('2024-01-01');
$mut = DateTime::createFromInterface($imm);
$mut->modify('+1 day');
echo $imm->format('Y-m-d') . ',' . $mut->format('Y-m-d'), "\n";

__vybe_check(ob_get_clean(), "2024-01-01,2024-01-02");
