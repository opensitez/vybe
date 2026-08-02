<?php
// vybe-test: php/datetimezone_get_transitions/datetimezone_get_transitions_reverse_range
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

$tz = new DateTimeZone('Europe/Paris');
$end = strtotime('2024-01-01');
$start = strtotime('2024-06-01');
$transitions = $tz->getTransitions($start, $end);
echo is_array($transitions) ? 'array' : 'na';

__vybe_check(ob_get_clean(), "array");
