<?php
// vybe-test: php/datetime/microtime_integer_and_float_runtime
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

echo is_float(microtime(true)) ? 'float' : 'int';
$micro = microtime();
echo str_contains($micro, ' ') ? 'sp' : 'ns';
echo strpos($micro, '.') !== false ? '|dot' : '|nodot';

__vybe_check(ob_get_clean(), "floatsp|dot");
