<?php
// vybe-test: php/arrays/array_map_with_keys_not_modified
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

$map = ['a' => 1, 'b' => 2];
$out = array_map(fn($v) => $v * 2, $map);
$first = key($out);
echo $first . '|' . $out['a'];

__vybe_check(ob_get_clean(), "a|2");
