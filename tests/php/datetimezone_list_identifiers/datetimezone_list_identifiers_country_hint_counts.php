<?php
// vybe-test: php/datetimezone_list_identifiers/datetimezone_list_identifiers_country_hint_counts
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

$zones = DateTimeZone::listIdentifiers(DateTimeZone::PER_COUNTRY, 'US');
echo is_array($zones) ? (count($zones) > 0 ? "has_us" : "no_us") : "bad";

__vybe_check(ob_get_clean(), "has_us");
