<?php
// vybe-test: php/datetime_parsing/strtotime_with_rfc3339
// origin: languages/php/tests/php/test_datetime_parsing.rs

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

$ts = strtotime('2024-03-01T13:45:00+00:00');
echo $ts === false ? 'bad' : 'ok';
echo '|';
echo date('Y-m-d H:i', $ts);

__vybe_check(ob_get_clean(), "ok|2024-03-01 13:45");
