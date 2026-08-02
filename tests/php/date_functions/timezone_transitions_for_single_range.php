<?php
// vybe-test: php/date_functions/timezone_transitions_for_single_range
// origin: languages/php/tests/php/test_date_functions.rs

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

$tz = new DateTimeZone('America/New_York');
$transitions = $tz->getTransitions(1704067200, 1706745600);
echo is_array($transitions) ? 'array' : 'bad';
echo '|';
echo count($transitions) > 1 ? 'many' : 'few';

__vybe_check(ob_get_clean(), "array|many");
