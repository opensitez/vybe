<?php
// vybe-test: php/php_datetime_timezones/php_datetime_timezone_abbreviations_and_transition_keys
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

$tz = new DateTimeZone('Europe/Berlin');
$abbrevs = DateTimeZone::listAbbreviations();
echo array_key_exists('ce', $abbrevs) ? 'ce' : 'no';
$transitions = $tz->getTransitions(strtotime('2024-01-01'), strtotime('2024-07-01'));
echo count($transitions) > 0 ? 'has' : 'none';

__vybe_check(ob_get_clean(), "nohas");
