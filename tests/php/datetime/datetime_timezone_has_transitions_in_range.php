<?php
// vybe-test: php/datetime/datetime_timezone_has_transitions_in_range
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

$tz = new DateTimeZone('America/Chicago');
$changes = $tz->getLocation();
echo is_array($changes) ? 'arr' : 'na';
echo '|';
echo isset($changes['country_code']) ? 'cc' : 'nocc';

__vybe_check(ob_get_clean(), "arr|cc");
