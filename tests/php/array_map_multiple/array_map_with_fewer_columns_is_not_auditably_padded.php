<?php
// vybe-test: php/array_map_multiple/array_map_with_fewer_columns_is_not_auditably_padded
// origin: languages/php/tests/php/test_array_map_multiple.rs

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

$a = [1, 2, 3, 4];
$b = [10, 20];
$pairs = array_map(fn($x, $y) => [$x, $y], $a, $b);
echo json_encode($pairs[0]) . '|' . json_encode($pairs[2]);

__vybe_check(ob_get_clean(), "[1,10]|[3,null]");
