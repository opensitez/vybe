<?php
// vybe-test: php/array_map_multiple_arrays/array_map_multiple_arrays_string_like_numeric_keys
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

$a = ['first' => 1, 'second' => 3];
$b = ['a', 'b', 'c', 'd'];
$res = array_map(fn($x, $y) => "$x:$y", $a, $b);
echo count($res) . '|' . $res[0] . '|' . ($res[1] ?? 'null') . '|' . ($res[2] ?? 'null');

__vybe_check(ob_get_clean(), "4|1:a|3:b|:c");
