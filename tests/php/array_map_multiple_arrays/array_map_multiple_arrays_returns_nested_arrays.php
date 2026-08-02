<?php
// vybe-test: php/array_map_multiple_arrays/array_map_multiple_arrays_returns_nested_arrays
// origin: languages/php/tests/php/test_array_map_multiple_arrays.rs

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

$a = [1, 2];
$b = ['u', 'v'];
$res = array_map(fn($x, $y) => ['v' => $x, 'k' => $y], $a, $b);
echo json_encode($res[0]) . "|" . json_encode($res[1]);

__vybe_check(ob_get_clean(), "{\"v\":1,\"k\":\"u\"}|{\"v\":2,\"k\":\"v\"}");
