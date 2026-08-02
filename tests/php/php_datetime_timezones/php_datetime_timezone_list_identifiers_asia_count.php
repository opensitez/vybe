<?php
// vybe-test: php/php_datetime_timezones/php_datetime_timezone_list_identifiers_asia_count
// origin: languages/php/tests/php/test_php_datetime_timezones.rs

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

$zones = DateTimeZone::listIdentifiers(DateTimeZone::ASIA);
echo is_array($zones) ? 'array' : 'bad';
echo '|';
echo in_array('Asia/Tokyo', $zones, true) ? 'has_tokyo' : 'no_tokyo';

__vybe_check(ob_get_clean(), "array|has_tokyo");
