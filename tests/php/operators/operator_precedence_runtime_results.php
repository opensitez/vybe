<?php
// vybe-test: php/operators/operator_precedence_runtime_results
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

echo (1 + 2) * 3;
echo (4 + 6) / 2;
echo 2 ** 3 ** 2;
echo (2 ** 3) ** 2;
echo 1 + 2 * 3;
echo 1 + 2 << 2;
echo (1 + 2) << 2;
echo 8 >> 1 + 1;
echo (8 >> 1) + 1;
echo 3 + 4 * 2 < 20 && 3 < 4;

__vybe_check(ob_get_clean(), "955126471212251");
