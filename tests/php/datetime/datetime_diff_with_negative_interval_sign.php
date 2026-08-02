<?php
// vybe-test: php/datetime/datetime_diff_with_negative_interval_sign
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

$start = new DateTime('2024-06-10');
$end = new DateTime('2024-06-01');
$diff = $start->diff($end);
echo $diff->invert ? 'neg' : 'pos';
echo '|';
echo $diff->days;

__vybe_check(ob_get_clean(), "neg|9");
