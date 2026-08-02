<?php
// vybe-test: php/php_date_time_zone_identifiers_transitions/test_php_datetimezone_transitions_minmax_range
// origin: languages/php/tests/php/test_php_date_time_zone_identifiers_transitions.rs

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

$tz = new DateTimeZone("America/Sao_Paulo");
$transitions = $tz->getTransitions(strtotime("2024-01-01"), strtotime("2024-12-31"));
echo is_array($transitions) ? 'array' : 'no';
echo '|';
echo count($transitions) > 0 ? 'has' : 'no';

__vybe_check(ob_get_clean(), "array|has");
