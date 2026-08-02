<?php
// vybe-test: php/strtotime_relative_keywords/strtotime_relative_keywords
// origin: languages/php/tests/php/test_strtotime_relative_keywords.rs

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

date_default_timezone_set('UTC');
$base = strtotime('2020-01-15 10:00:00');
echo date('Y-m-d H:i:s', strtotime('+1 day', $base)) . "|";
echo date('Y-m-d H:i:s', strtotime('next monday', $base));

__vybe_check(ob_get_clean(), "2020-01-16 10:00:00|2020-01-20 00:00:00");
