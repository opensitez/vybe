<?php
// vybe-test: php/literals/test_php_float_and_nan_in_literals
// origin: languages/php/tests/php/test_literals.rs

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

echo round(1.25 + 0.5, 1);
echo '|';
echo is_float(1.0 / 2);
echo '|';
echo (NAN === NAN) ? 'nan' : 'not';

__vybe_check(ob_get_clean(), "1.8|1|not");
