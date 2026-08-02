<?php
// vybe-test: php/dateinterval_create_from_datestring/dateinterval_create_from_datestring_invalid
// origin: languages/php/tests/php/test_dateinterval_create_from_datestring.rs

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

$i = DateInterval::createFromDateString('not a valid duration token');
echo $i === false ? 'false' : 'notfalse';

__vybe_check(ob_get_clean(), "notfalse");
