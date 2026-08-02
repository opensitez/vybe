<?php
// vybe-test: php/date_functions/datetimeimmutable_add_returns_new_instance
// origin: languages/php/tests/php/test_date_functions.rs

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

$orig = new DateTimeImmutable('2024-01-01');
$new = $orig->add(new DateInterval('P1D'));
echo $orig->format('d') . ':' . $new->format('d');

__vybe_check(ob_get_clean(), "01:02");
