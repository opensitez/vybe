<?php
// vybe-test: php/mixed_programs/matrix_multiplication
// origin: languages/php/tests/php/test_mixed_programs.rs

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

function matmul(array $A, array $B): array {
    $result = [];
    for ($i = 0; $i < count($A); $i++) {
        for ($j = 0; $j < count($B[0]); $j++) {
            $result[$i][$j] = 0;
            for ($k = 0; $k < count($B); $k++) {
                $result[$i][$j] += $A[$i][$k] * $B[$k][$j];
            }
        }
    }
    return $result;
}
$A = [[1,2],[3,4]];
$B = [[5,6],[7,8]];
$C = matmul($A, $B);
echo $C[0][0] . ',' . $C[0][1] . ',' . $C[1][0] . ',' . $C[1][1];

__vybe_check(ob_get_clean(), "19,22,43,50");
