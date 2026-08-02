<?php
// vybe-test: php/datetime/datetime_timezone_abbreviations_snapshot
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

$abbr = DateTimeZone::listAbbreviations();
echo is_array($abbr) ? 'arr' : 'na';
echo '|';
echo isset($abbr['est']) ? 'has' : 'missing';

__vybe_check(ob_get_clean(), "arr|has");
