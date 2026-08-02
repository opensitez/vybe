<?php
// vybe-test: php/php_datetime_immutable_interval_period/test_php_dateinterval_component_boundaries
// origin: languages/php/tests/php/test_php_datetime_immutable_interval_period.rs

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

$i = new DateInterval('P1Y2M3DT4H5M6S');
echo $i->y . $i->m . $i->d . $i->h . $i->i . $i->s;

__vybe_check(ob_get_clean(), "123456");
