<?php
// vybe-test: php/datetimezone_get_transitions/datetimezone_get_transitions
// origin: languages/php/tests/php/test_datetimezone_get_transitions.rs

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

$tz = new DateTimeZone('Europe/London');
$transitions = $tz->getTransitions(
    strtotime('2020-03-25'),
    strtotime('2020-04-05')
);

// We expect a DST transition on 2020-03-29
echo count($transitions) . "|";
if (count($transitions) > 1) {
    echo $transitions[1]['isdst'] ? 'DST' : 'STD';
}

__vybe_check(ob_get_clean(), "2|DST");
