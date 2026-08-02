<?php
// vybe-test: php/operators/bitwise_and_shift_precedence_runtime
// origin: languages/php/tests/php/test_operators.rs

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

echo (3 & 6 | 9) . '|';
echo ((3 & 6) | 9) . '|';
echo (3 & (6 | 9)) . '|';
echo (1 << 2 | 3) . '|';
echo (1 | 2 << 3);

__vybe_check(ob_get_clean(), "11|11|3|7|17");
