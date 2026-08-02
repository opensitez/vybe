<?php
// vybe-test: php/datetime/test_datetime_immutable_modify_does_not_mutate_original
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

$a = new DateTimeImmutable('2024-01-01');
$b = $a->modify('+10 days');
echo $a->format('Y-m-d');
echo $b->format('Y-m-d');

__vybe_check(ob_get_clean(), "2024-01-012024-01-11");
