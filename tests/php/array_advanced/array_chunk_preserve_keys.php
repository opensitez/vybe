<?php
// vybe-test: php/array_advanced/array_chunk_preserve_keys
// origin: languages/php/tests/php/test_array_advanced.rs

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

$a = ["a" => 1, "b" => 2, "c" => 3, "d" => 4];
$chunks = array_chunk($a, 3, true);
echo count($chunks);
echo implode(",", array_keys($chunks[0]));

__vybe_check(ob_get_clean(), "2a,b,c");
