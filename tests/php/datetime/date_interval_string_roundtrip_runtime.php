<?php
// vybe-test: php/datetime/date_interval_string_roundtrip_runtime
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

$i = new DateInterval('P1DT2H30M15S');
echo $i->d;
echo $i->h;
echo $i->i;
echo $i->s;

__vybe_check(ob_get_clean(), "123015");
