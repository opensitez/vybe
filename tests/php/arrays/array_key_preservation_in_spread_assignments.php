<?php
// vybe-test: php/arrays/array_key_preservation_in_spread_assignments
// origin: languages/php/tests/php/test_arrays.rs

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

$left = ['a' => 1, 2 => 2];
$right = [...$left, 'b' => 3, 4 => 4];
ksort($right);
echo json_encode($right);

__vybe_check(ob_get_clean(), "{\"0\":2,\"4\":4,\"a\":1,\"b\":3}");
