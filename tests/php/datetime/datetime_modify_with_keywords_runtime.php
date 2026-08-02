<?php
// vybe-test: php/datetime/datetime_modify_with_keywords_runtime
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

$dt = new DateTime('2024-01-31');
$dt->modify('first day of next month');
echo $dt->format('Y-m-d');
echo '|';
$dt->modify('last day of last month');
echo $dt->format('Y-m-d');

__vybe_check(ob_get_clean(), "2024-02-01|2024-01-31");
