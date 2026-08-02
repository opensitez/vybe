<?php
// vybe-test: php/arrays/array_splice_returns_removed_indexed_from_zero
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

$a = ['k' => 'v1', 'm' => 'v2', 'n' => 'v3'];
$r = array_splice($a, 1, 1, ['x' => 'vv']);
echo isset($r[0]) ? $r[0] : 'none';
echo ':' . count($r);

__vybe_check(ob_get_clean(), "v2:1");
