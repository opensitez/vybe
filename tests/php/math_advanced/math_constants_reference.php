<?php
// vybe-test: php/math_advanced/math_constants_reference
// origin: languages/php/tests/php/test_math_advanced.rs

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

echo (round(M_2_PI, 8) == 0.63661977 ? 'ok' : 'bad') . "\n";
echo (round(M_SQRT3, 8) == 1.73205081 ? 'ok' : 'bad') . "\n";
echo (M_LOG10E > 0 && M_LN10 > 0 ? 'ok' : 'bad') . "\n";
echo (M_LN2 < M_LN10 ? 'ok' : 'bad') . "\n";

__vybe_check(ob_get_clean(), "ok\nok\nok\nok");
