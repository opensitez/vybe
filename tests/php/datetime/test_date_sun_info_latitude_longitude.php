<?php
// vybe-test: php/datetime/test_date_sun_info_latitude_longitude
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

$sun = date_sun_info(mktime(0, 0, 0, 6, 21, 2024), 51.5074, -0.1278);
echo is_array($sun) && isset($sun['sunrise']) ? 'sun_info_ok' : 'err';

__vybe_check(ob_get_clean(), "sun_info_ok");
