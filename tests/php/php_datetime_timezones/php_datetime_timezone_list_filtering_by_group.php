<?php
// vybe-test: php/php_datetime_timezones/php_datetime_timezone_list_filtering_by_group
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

$tzs = DateTimeZone::listIdentifiers(DateTimeZone::AMERICA);
echo is_array($tzs) ? 'array' : 'no';
echo '|';
echo in_array('America/New_York', $tzs, true) ? 'ny' : 'no';

__vybe_check(ob_get_clean(), "array|ny");
