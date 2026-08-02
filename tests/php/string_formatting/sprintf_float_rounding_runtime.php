<?php
// vybe-test: php/string_formatting/sprintf_float_rounding_runtime
// origin: languages/php/tests/php/test_string_formatting.rs

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

echo sprintf('%.0f', 2.5);
echo "\n";
echo sprintf('%.2e', 1.2);
echo "\n";
echo sprintf('%.3f', -1.2345);

__vybe_check(ob_get_clean(), "3|1.20e+00|-1.235");
