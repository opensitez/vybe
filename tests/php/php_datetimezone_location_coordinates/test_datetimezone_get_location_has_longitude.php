<?php
// vybe-test: php/php_datetimezone_location_coordinates/test_datetimezone_get_location_has_longitude
// origin: languages/php/tests/php/test_php_datetimezone_location_coordinates.rs

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

$tz = new DateTimeZone('Asia/Tokyo');
$loc = $tz->getLocation();
echo is_numeric($loc['longitude']) ? 'lon' : 'nolon';

__vybe_check(ob_get_clean(), "lon");
