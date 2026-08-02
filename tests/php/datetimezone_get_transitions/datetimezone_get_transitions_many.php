<?php
// vybe-test: php/datetimezone_get_transitions/datetimezone_get_transitions_many
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

$tz = new DateTimeZone('America/Sao_Paulo');
$start = strtotime('2023-01-01');
$end = strtotime('2023-12-31');
$transitions = $tz->getTransitions($start, $end);
echo is_array($transitions) ? 'ok' : 'bad';
echo '|';
echo count($transitions) > 1 ? 'many' : 'few';

__vybe_check(ob_get_clean(), "ok|many");
