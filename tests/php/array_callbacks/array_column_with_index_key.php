<?php
// vybe-test: php/array_callbacks/array_column_with_index_key
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

$rows = [['id' => 10, 'v' => 'x'], ['id' => 20, 'v' => 'y']];
$map = array_column($rows, 'v', 'id');
echo $map[10] . $map[20];

__vybe_check(ob_get_clean(), "xy");
