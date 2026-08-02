<?php
// vybe-test: php/datetime/test_checkdate_leap_year
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

echo checkdate(2, 29, 2024) ? 'leap_ok' : 'err';
echo checkdate(2, 29, 2023) ? 'err' : ' non_leap_ok';

__vybe_check(ob_get_clean(), "leap_ok non_leap_ok");
