<?php
// vybe-test: php/list_destructuring/list_matrix_row_extraction
// origin: languages/php/tests/php/test_list_destructuring.rs

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

$matrix = [[1,2,3],[4,5,6],[7,8,9]];
[[$a,$b,$c], , [$g,$h,$i]] = $matrix;
echo $a + $i;

__vybe_check(ob_get_clean(), "10");
