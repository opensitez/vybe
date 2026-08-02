<?php
// vybe-test: php/array_functions/array_replace_recursive_merges_nested
// origin: languages/php/tests/php/test_array_functions.rs

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

$base = ['cfg' => ['a' => 1, 'b' => 2]];
$over = ['cfg' => ['b' => 9]];
$r = array_replace_recursive($base, $over);
echo $r['cfg']['a'] . ':' . $r['cfg']['b'];

__vybe_check(ob_get_clean(), "1:9");
