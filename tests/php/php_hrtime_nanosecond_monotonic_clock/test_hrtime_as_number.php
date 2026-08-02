<?php
// vybe-test: php/php_hrtime_nanosecond_monotonic_clock/test_hrtime_as_number
// origin: languages/php/tests/php/test_php_hrtime_nanosecond_monotonic_clock.rs

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

$t = hrtime(true);
echo (is_int($t) || is_float($t)) && $t > 0 ? 'number_ok' : 'err', "\n";

__vybe_check(ob_get_clean(), "number_ok");
