<?php
// vybe-test: php/php_datetime_immutable_interval_period/test_php_datetimeimmutable_chain_does_not_mutate
// origin: languages/php/tests/php/test_php_datetime_immutable_interval_period.rs

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
$b = $a->add(new DateInterval('P1D'))->modify('+1 day')->setTime(9, 0);
echo $a->format('Y-m-d') . '|' . $b->format('Y-m-d H:i:s');

__vybe_check(ob_get_clean(), "2024-01-01|2024-01-03 09:00:00");
