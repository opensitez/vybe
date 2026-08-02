<?php
// vybe-test: php/php_array_chunk_combine_count_values/test_php_array_intersect_assoc_preserves_key_order
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

$a1 = ["x" => 1, "y" => 2, "z" => 3];
$a2 = ["y" => 2, "x" => 1];
$i = array_intersect_assoc($a1, $a2);
echo implode(",", array_keys($i)) . ":" . implode(",", $i);

__vybe_check(ob_get_clean(), "x,y:1,2");
