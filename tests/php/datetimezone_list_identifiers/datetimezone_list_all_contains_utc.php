<?php
// vybe-test: php/datetimezone_list_identifiers/datetimezone_list_all_contains_utc
// origin: languages/php/tests/php/test_datetimezone_list_identifiers.rs

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

$zones = DateTimeZone::listIdentifiers();
echo in_array('UTC', $zones) ? "utc" : "no_utc";

__vybe_check(ob_get_clean(), "utc");
