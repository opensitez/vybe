<?php
// vybe-test: php/datetime/datetime_timezone_transition_runtime
// origin: languages/php/tests/php/test_datetime.rs

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
$trans = $tz->getTransitions(strtotime('2024-01-01'), strtotime('2024-12-31'));
echo is_array($trans) ? 'arr' : 'na';
echo '|';
echo count($trans) > 1 ? 'many' : 'few';

__vybe_check(ob_get_clean(), "arr|many");
