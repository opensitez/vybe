<?php
// vybe-test: php/array_callbacks/array_map_truncates_to_shortest_input
// origin: languages/php/tests/php/test_array_callbacks.rs

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

$a = [1, 2, 3];
$b = [10, 20];
$mapped = array_map(fn($x, $y) => $x + $y, $a, $b);
echo json_encode($mapped);

__vybe_check(ob_get_clean(), "[11,22,3]");
