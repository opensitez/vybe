<?php
// vybe-test: php/php_array_chunk_combine_count_values/test_php_array_diff_assoc_and_intersect_assoc
// origin: languages/php/tests/php/test_php_array_chunk_combine_count_values.rs

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

$a1 = ["a" => "green", "b" => "brown", "c" => "blue", "red"];
$a2 = ["a" => "green", "yellow", "red"];

$diff = array_diff_assoc($a1, $a2);
$intersect = array_intersect_assoc($a1, $a2);

echo "DiffCount=" . count($diff) . " IntersectCount=" . count($intersect);

__vybe_check(ob_get_clean(), "DiffCount=3 IntersectCount=1");
