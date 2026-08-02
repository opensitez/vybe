<?php
// vybe-test: php/datetimezone_get_transitions/datetimezone_dst_interval_snapshot
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

$tz = new DateTimeZone('America/Los_Angeles');
$start = strtotime('2021-03-01 00:00:00');
$end = strtotime('2021-04-01 00:00:00');
$transitions = $tz->getTransitions($start, $end);
echo count($transitions) >= 1 ? 'has' : 'none';
if (count($transitions) > 0 && isset($transitions[0]['isdst'])) {
    echo '|';
    echo $transitions[0]['isdst'] ? 'dst' : 'std';
}

__vybe_check(ob_get_clean(), "has|std");
